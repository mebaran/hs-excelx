{-# LANGUAGE QuasiQuotes, OverloadedStrings, FlexibleContexts, FlexibleInstances #-}

module Data.Excel where

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
          | BlankCell Position deriving Show

type SheetIndex = [(T.Text, FilePath)]
data Excelx = Excelx {sharedStrings :: [T.Text], sheetIndex :: SheetIndex, archive :: Archive}

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
                   _ -> Just (fromCell c)
    
a1form :: Position -> T.Text
a1form (A1 t) = t
a1form (R1C1 i j) = T.toUpper . T.pack $ alpha (max j 1) ++ show (max i 1)
  where 
    alphabet = map reverse [c : s | s <- "" : alphabet, c <- ['a' .. 'z']]
    alpha idx = alphabet !! (idx - 1)
          
findXMLEntry :: FilePath -> Archive -> Maybe Cursor
findXMLEntry path ar =
  fmap (fromDocument . parseLBS_ def . fromEntry) $ findEntryByPath path ar
  
extractSheetIndex :: Archive -> Maybe [(T.Text, String)]
extractSheetIndex ar = do
  sheetXml <- findXMLEntry "xl/workbook.xml" ar
  relXml <- findXMLEntry "xl/_rels/workbook.xml.rels" ar
  return $ do
    let relationships = relXml $.// laxElement "Relationship"
        sheetListings =  sheetXml $.// laxElement "sheet"
    sheetListingEntry <- sheetListings 
    relationship <- relationships
    guard $ ((laxAttribute "Id") relationship) == ((laxAttribute "id") sheetListingEntry)
    sheetName <- (laxAttribute "Name") sheetListingEntry
    target <- (laxAttribute "Target") relationship
    return (sheetName, T.unpack $ mappend "xl/" target)

extractSharedStrings :: Archive -> Maybe [T.Text]
extractSharedStrings ar = do
  sharedStringXml <- findXMLEntry "xl/sharedStrings.xml" ar
  return $ sharedStringXml $.// laxElement "sst" &/ laxElement "si" &.// content

parseCell :: Cursor -> Position -> Excelx -> Maybe Cell
parseCell xmlSheet cellpos xlsx = 
  let at = a1form cellpos 
      cursor = xmlSheet $// laxElement "c" >=> attributeIs "r" at in
  case listToMaybe cursor of
    Nothing -> Nothing
    Just c -> msum $ map (\p -> p xlsx at c) [parseFormulaCell,
                                              parseNumericCell,
                                              parseTextCell,
                                              parseBlankCell]
      where
        cellContents tag xml = xml $.// laxElement tag &.// content    

        parseFormulaCell xlsx at xml = do
          formula <- listToMaybe $ cellContents "f" xml
          return $ FormulaCell (A1 at) formula $ 
            fromMaybe (BlankCell (A1 at)) $ msum [parseNumericCell xlsx at xml,
                                                  parseTextCell xlsx at xml,
                                                  parseBlankCell xlsx at xml]        
        parseTextCell xlsx at xml = do
          textcell <- listToMaybe $ xml $.// attributeIs "t" "s" 
          val <- listToMaybe $ cellContents "v" textcell
          return $ TextCell (A1 at) (sharedStrings xlsx !! read (T.unpack val))
     
        parseNumericCell xlsx at xml = do
          v <- listToMaybe $ cellContents "v" xml
          return $ NumericCell (A1 at) (read $ T.unpack v)
            
        parseBlankCell xlsx at xml = return $ BlankCell (A1 at)

extractSheet :: T.Text -> Excelx -> Maybe Cursor
extractSheet sheetname xlsx = do
  sheetPath <- lookup sheetname $ sheetIndex xlsx
  findXMLEntry sheetPath $ archive xlsx
  
toExcelx :: BS.ByteString -> Maybe Excelx
toExcelx bytes = do
  let ar = toArchive bytes
  sharedStringList <- extractSharedStrings ar
  sheetIdx <- extractSheetIndex ar
  return $ Excelx sharedStringList sheetIdx ar

openExcelx :: FilePath -> IO (Maybe Excelx)
openExcelx f = do
  ar <- BS.readFile f
  return $ toExcelx ar

sheet :: MonadReader Excelx m => T.Text -> m (Maybe Cursor)
sheet name = liftM (extractSheet name) ask

cell :: MonadReader Excelx m => Cursor -> Position -> m (Maybe Cell)
cell sheetCur cellpos = liftM (parseCell sheetCur cellpos) ask

column :: MonadReader Excelx m => Cursor -> Int -> m [Maybe Cell]
column sheetCur colidx = mapM (cell sheetCur) col where col = map (flip R1C1 colidx) [1..]

runExcel :: b -> Reader b c -> c
runExcel = flip runReader
