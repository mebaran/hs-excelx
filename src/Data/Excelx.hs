{-# LANGUAGE QuasiQuotes, OverloadedStrings, FlexibleContexts, FlexibleInstances #-}

module Data.Excelx(openExcelx, toExcelx,
                   sheet, cell, row, rows, column,
                   fromCell) where

import Data.Maybe
import Data.Monoid

import Data.Time.LocalTime
import Data.Time.Calendar

import Control.Monad
import Control.Monad.Reader

import qualified Data.Text as T
import qualified Data.ByteString.Lazy as BS 
import Codec.Archive.Zip

import Text.XML as XML
import Text.XML.Cursor 

type Formula = T.Text
data Position = R1C1 Int Int | A1 T.Text deriving Show
data Cell = NumericCell Position Double
          | TextCell Position T.Text
          | FormulaCell Position Formula Cell
          | BlankCell Position 
          | NoCell Position deriving Show

valueCell :: Cell -> Cell
valueCell (FormulaCell _ _ valuecell) = valuecell
valueCell valuecell = valuecell

class FromCell a where
    fromCell :: Cell -> a

instance FromCell Double where
    fromCell c = case valueCell c of
                   NumericCell _ d -> d
                   _ -> 0.0

instance FromCell Integer where
    fromCell c = floor (fromCell c :: Double)

instance FromCell T.Text where
    fromCell c = case valueCell c of
                   TextCell _ t -> t
                   NumericCell _ t -> T.pack $ show t
                   _ -> ""

instance FromCell String where
    fromCell c = T.unpack $ fromCell c

instance FromCell Day where
    fromCell c = addDays (fromCell c) (fromGregorian 1899 12 30)

instance FromCell TimeOfDay where
    fromCell c = dayFractionToTimeOfDay part
        where
          d = fromCell c :: Double
          part = toRational $ d - (fromInteger $ floor d)

instance FromCell LocalTime where
    fromCell c = LocalTime (fromCell c) (fromCell c)

instance FromCell a => FromCell (Maybe a) where
    fromCell c = case valueCell c of
                   BlankCell _ -> Nothing
                   NoCell _ -> Nothing
                   _ -> Just (fromCell c)

instance FromCell a => FromCell (Either Position a) where
    fromCell c = case valueCell c of
                   NoCell pos -> Left pos
                   _ -> Right $ fromCell c

catCells :: [Cell] -> [Cell]
catCells cells = filter noCell cells
    where noCell c = case c of
              NoCell _ -> True
              _ -> False

a1form :: Position -> T.Text
a1form (A1 t) = t
a1form (R1C1 i j) = T.toUpper . T.pack $ alpha (max j 1) ++ show (max i 1)
  where 
    alphabet = map reverse [c : s | s <- "" : alphabet, c <- ['a' .. 'z']]
    alpha idx = alphabet !! (idx - 1)

type SharedStringIndex = [T.Text]
type SheetIndex = [(T.Text, FilePath)]
data Excelx = Excelx {archive :: Archive, sharedStrings :: SharedStringIndex, sheetIndex :: SheetIndex}
          
findXMLEntry :: FilePath -> Archive -> Maybe Cursor
findXMLEntry path ar = do
  entry <- findEntryByPath path ar
  case parseLBS def (fromEntry entry) of
    Left _ -> fail $ "Invalid entry: " ++ path 
    Right xml -> return $ fromDocument xml
  
extractSheetIndex :: Archive -> Maybe [(T.Text, String)]
extractSheetIndex ar = do
  sheetXml <- findXMLEntry "xl/workbook.xml" ar
  relXml <- findXMLEntry "xl/_rels/workbook.xml.rels" ar
  return $ do
    let relationships = relXml $.// laxElement "Relationship"
        sheetListings =  sheetXml $.// laxElement "sheet"
    sheetListingEntry <- sheetListings 
    relationship <- relationships
    guard $ (laxAttribute "Id") relationship == (laxAttribute "id") sheetListingEntry
    sheetName <- (laxAttribute "Name") sheetListingEntry
    target <- (laxAttribute "Target") relationship
    return (sheetName, T.unpack $ mappend "xl/" target)

extractSharedStrings :: Archive -> Maybe [T.Text]
extractSharedStrings ar = do
  sharedStringXml <- findXMLEntry "xl/sharedStrings.xml" ar
  let entries = sharedStringXml $.// laxElement "si"
  return $ map mconcat $ map (\e -> e $// laxElement "t" &/ content) entries

parseCell :: Excelx -> Cursor -> Position -> Cell
parseCell xlsx xmlSheet cellpos = 
  let at = a1form cellpos 
      cursor = xmlSheet $// laxElement "c" >=> attributeIs "r" at in
  case listToMaybe cursor of
    Nothing -> NoCell cellpos
    Just c -> 
        let formulaCell = do
              formula <- listToMaybe $ cellContents "f" c
              let valcell = fromMaybe (BlankCell cellpos) $ msum [textCell, numericCell, blankCell]
              return $ FormulaCell cellpos formula valcell
            textCell = do
              textcell <- listToMaybe $ c $.// attributeIs "t" "s"
              val <- listToMaybe $ cellContents "v" textcell
              return $ TextCell cellpos (sharedStrings xlsx !! (read (T.unpack val)))
            numericCell = do
              v <- listToMaybe $ cellContents "v" c
              return $ NumericCell cellpos (read $ T.unpack v)
            blankCell = return $ BlankCell cellpos
        in fromJust $ msum [formulaCell, textCell, numericCell, blankCell]
      where
        cellContents tag xml = xml $.// laxElement tag &.// content

parseRow :: Excelx -> Cursor -> Int -> [Cell]
parseRow xlsx sheetCur rownum = 
    let rowxml = sheetCur $// laxElement "row" >=> attributeIs "r" (T.pack $ show rownum)
        addrs = concatMap (\rx -> rx $// laxElement "c" >=> attribute "r") rowxml
    in map (parseCell xlsx sheetCur) (map A1 addrs)

maxRow :: Cursor -> Int
maxRow sheetxml = foldr max 0 $ map (read . T.unpack) $ sheetxml $.// laxElement "row" >=> laxAttribute "r"

extractSheet :: T.Text -> Excelx -> Maybe Cursor
extractSheet sheetname xlsx = do
  sheetPath <- lookup sheetname $ sheetIndex xlsx
  findXMLEntry sheetPath $ archive xlsx
  
toExcelx :: BS.ByteString -> Maybe Excelx
toExcelx bytes = do
  let ar = toArchive bytes
  sharedStringList <- extractSharedStrings ar
  sheetIdx <- extractSheetIndex ar
  return $ Excelx ar sharedStringList sheetIdx

openExcelx :: FilePath -> IO (Maybe Excelx)
openExcelx f = do
  ar <- BS.readFile f
  return $ toExcelx ar

type Sheet = (Excelx, Cursor)
type Row = [Cell]

sheet :: MonadReader Excelx m => T.Text -> m (Maybe Sheet)
sheet name = do
  xlsx <- ask
  let sheetx = extractSheet name xlsx
  return $ fmap (\sheetxml -> (xlsx, sheetxml)) sheetx

row :: MonadReader Sheet m => Int -> m Row
row num = do
  (xlsx, sheetx) <- ask
  return $ parseRow xlsx sheetx num

rows :: MonadReader Sheet m => m [Row]
rows = do
  (_, sheetx) <- ask
  mapM row [1 .. maxRow sheetx]

cell :: MonadReader Sheet m => Position -> m (Cell)  
cell cellpos = do
  (xlsx, sheetx) <- ask
  return $ parseCell xlsx sheetx cellpos

column :: MonadReader Sheet m => Int -> m [Cell]
column colidx = do
  (_, sheetx) <- ask
  liftM catCells $ mapM cell $ map (flip R1C1 colidx) [1 .. maxRow sheetx]

inExcel :: b -> Reader b c -> c
inExcel = flip runReader

inSheet :: MonadReader Excelx m => T.Text -> Reader Sheet a -> m (Maybe a)
inSheet name action = do
  maybeSheet <- sheet name
  case maybeSheet of
    Just sheetx -> return $ Just $ runReader action sheetx
    Nothing -> return Nothing

inExcelSheet :: Excelx -> T.Text -> Reader Sheet a -> Maybe a
inExcelSheet xlsx name action = inExcel xlsx $ inSheet name action
