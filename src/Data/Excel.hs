{-# LANGUAGE QuasiQuotes, OverloadedStrings #-}
module Data.Excel where

import Data.List
import Data.Maybe
import qualified Data.Map as M 
import Data.Monoid
import Data.Time.Calendar

import Control.Monad
import Control.Monad.Reader

import qualified Data.Text as T
import qualified Data.ByteString.Lazy as BS 
import Codec.Archive.Zip

import Text.XML as XML
import Text.XML.Cursor 
import Text.XML.Scraping (innerText, innerTextN)
import Text.XML.Selector.TH

type Formula = T.Text
data Position = R1C1 Int Int | A1 T.Text deriving Show
data Cell = NumericCell Position Double
          | DateCell Position Day 
          | TextCell Position T.Text 
          | FormulaCell Position Formula Cell 
          | BlankCell Position deriving Show

type SheetIndex = [(T.Text, FilePath)]
data Excelx = Excelx {sharedStrings :: [T.Text], sheetIndex :: SheetIndex, archive :: Archive}

a1form (A1 t) = t
a1form (R1C1 i j) = T.toUpper . T.pack $ (alpha (max i 1)) ++ show (max j 1)
  where 
    alphabet = map reverse [c : s | s <- "" : alphabet, c <- ['a' .. 'z']]
    alpha i = alphabet !! (i - 1)
          

findXMLEntry :: FilePath -> Archive -> Maybe Cursor
findXMLEntry path archive =
  fmap (fromDocument . (parseLBS_ def) . fromEntry) $ findEntryByPath path archive 

nodeElement :: Node -> Maybe Element
nodeElement (NodeElement element) = Just element
nodeElement _ = Nothing
  
extractSheetIndex :: Archive -> Maybe [(T.Text, String)]
extractSheetIndex archive = do
  sheetXml <- findXMLEntry "xl/workbook.xml" archive
  rels <- findXMLEntry "xl/_rels/workbook.xml.rels" archive
  let relationships = map (fmap elementAttributes) $ map (nodeElement . node) $ queryT [jq| Relationship |] rels 
      sheetListing = map (fmap elementAttributes) $ map (nodeElement . node) $ queryT [jq| sheet |] sheetXml
  return $ catMaybes $ do
    let relrid = fmap (M.lookup "Id")
        sheetrid = fmap (M.lookup rIdattrname)
    sheet <- sheetListing
    rel <- relationships
    guard $ (relrid rel) == (sheetrid sheet)
    return $ do
      jsheet <- sheet
      jrel <- rel
      sheetName <- M.lookup "name" jsheet
      target <- fmap (mappend "xl/") $ M.lookup "Target" jrel
      return (sheetName, T.unpack target)
  where 
    rIdattrname = Name {nameLocalName = "id", nameNamespace = Just "http://schemas.openxmlformats.org/officeDocument/2006/relationships", namePrefix = Just "r"}

extractSharedStrings :: Archive -> Maybe [T.Text]
extractSharedStrings archive = do
  sharedStringXml <- findXMLEntry "xl/sharedStrings.xml" archive
  return $ map (innerTextN . node) $ queryT [jq| sst si |] sharedStringXml
  
toExcelx :: BS.ByteString -> Maybe Excelx
toExcelx bytes = do
  let archive = toArchive bytes
  sharedStringList <- extractSharedStrings archive
  sheetIndex <- extractSheetIndex archive
  return $ Excelx sharedStringList sheetIndex archive
  
openExcelx :: FilePath -> IO (Maybe Excelx)
openExcelx f = do
  archive <- BS.readFile f
  return $ toExcelx archive

parseCell cellpos xmlSheet xlsx = 
  let at = a1form cellpos 
      cursor = xmlSheet $// (laxElement "c") >=> attributeIs "r" at in
  case listToMaybe cursor of
    Nothing -> (Nothing,Nothing)
    Just c -> (Just c, msum $ map (\p -> p xlsx at c) [parseFormulaCell,
                                              parseDateCell,
                                              parseNumericCell,
                                              parseTextCell,
                                              parseBlankCell])
      where
        cellContents tag xml = xml $.// laxElement tag &.// content
        
        parseFormulaCell xlsx at xml = do
          formula <- listToMaybe $ cellContents "f" xml
          return $ FormulaCell (A1 at) formula $ 
            fromMaybe (BlankCell (A1 at)) $ msum [parseDateCell xlsx at xml,
                                                  parseNumericCell xlsx at xml,
                                                  parseTextCell xlsx at xml,
                                                  parseBlankCell xlsx at xml]
            
        parseTextCell xlsx at xml = do
          cell <- listToMaybe $ xml $.// attributeIs "t" "s" 
          val <- listToMaybe $ cellContents "v" xml 
          return $ TextCell (A1 at) (sharedStrings xlsx !! (read $ T.unpack val))
            
        parseNumericCell xlsx at xml = do 
          v <- listToMaybe $ cellContents "v" xml
          return $ NumericCell (A1 at) (read $ T.unpack v)
            
        parseDateCell xlsx at xml = do
          style <- listToMaybe $ xml $.// attribute "s"
          if isDateStyle style then
            (do v <- listToMaybe $ cellContents "v" xml
                let date = addDays diff $ fromGregorian 1899 12 30
                    diff = read $ T.unpack v
                return $ DateCell (A1 at) date)
            else mzero
            where isDateStyle = (`elem` ["30", "31", "39", "14"])
            
        parseBlankCell xlsx at xml = return $ BlankCell (A1 at)

extractSheet sheetname xlsx = do
  sheetPath <- lookup sheetname $ sheetIndex xlsx
  findXMLEntry sheetPath $ archive xlsx

sheet name = liftM (extractSheet name) ask
cell sheet pos = liftM (parseCell sheet pos) ask

runExcel = flip runReader 