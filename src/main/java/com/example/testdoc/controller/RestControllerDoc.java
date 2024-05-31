package com.example.testdoc.controller;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/v1")
@Slf4j
public class RestControllerDoc {



    @GetMapping("/doc")
    public  String getDoc(){
        this.generateDoc();
        return  "document generated" ;
    }

    public  void  generateDoc(){

        Map<String,String> map = Map.of("nameValue","Houssem edd",
                "ageValue","22",
                "genreValue","male");
        // path were the original template is located
        String filePath = "D:\\genDoc\\template.docx";
        //path of the out put result
        String pathToSave = "D:\\genDoc\\out\\template.docx";
        try( OPCPackage fs = OPCPackage.open(new File(filePath));) {

            XWPFDocument doc = new XWPFDocument(new FileInputStream( filePath ));
            doc = replaceText(doc, map);
            saveWord(pathToSave, doc);
        } catch(FileNotFoundException e){
        log.error( "an exception has  bee thrown  "  + e.getMessage() ,e );
        }
        catch(IOException e){
            log.error( "an exception has  bee thrown  "  + e.getMessage() ,e );
        } catch (InvalidFormatException e) {

            log.error( "an exception has  bee thrown  "  + e.getMessage() ,e );
        }
    }


    private XWPFDocument replaceText(XWPFDocument doc, Map<String ,String>map) {
        for(Map.Entry<String,String> entry : map.entrySet()) {
            replaceTextInParagraphs( doc.getParagraphs(),entry.getKey(), entry.getValue() );
            for (XWPFTable tbl : doc.getTables()) {
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        replaceTextInParagraphs( cell.getParagraphs(), entry.getKey(), entry.getValue());
                    }
                }
            }
        }
        return doc;
    }

    private void replaceTextInParagraphs(List<XWPFParagraph> paragraphs, String originalText, String updatedText) {
        paragraphs.forEach(paragraph -> replaceTextInParagraph(paragraph, originalText, updatedText));
    }
    private void replaceTextInParagraph(XWPFParagraph paragraph, String originalText, String updatedText) {
        String paragraphText = paragraph.getParagraphText();
        if (paragraphText.contains(originalText)) {
            String updatedParagraphText = paragraphText.replace(originalText, updatedText);
            while (!paragraph.getRuns().isEmpty()) {
                paragraph.removeRun(0);
            }
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(updatedParagraphText);
        }
    }
    // saving result uder the specific path
    private static void saveWord(String pathToSave, XWPFDocument doc) throws IOException{


       try (OutputStream out = new FileOutputStream(new File( pathToSave ))) {
           doc.write( out );
       }

    }


}

