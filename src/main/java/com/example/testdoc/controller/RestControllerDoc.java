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
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
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

    public  void  generateDoc(Object request){

        Map<String,String> map = Map.of("nameValue","Houssem edd",
                "ageValue","22",
                "genreValue","male","SPECIFY THE NAME","Houssem Eddine Maissoudi");
        // path were the original template is located
        String filePath = "D:\\genDoc\\template.docx";
        //path of the out put result
        String pathToSave = "D:\\genDoc\\out\\template.docx";
        try( OPCPackage fs = OPCPackage.open(new File(filePath));) {

            XWPFDocument doc = new XWPFDocument(new FileInputStream( filePath ));
            doc = replaceText(doc, map );
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


    private XWPFDocument replaceText(XWPFDocument doc, Map<String ,String>map ) {
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

      /*  table = doc.createTable(2,3);
        table.setWidth( 300 );
        table.setColBandSize( 2 );
        table.getRow( 0 ).getCell( 0 ).setText( "NAME " );
        table.getRow( 0 ).getCell(  1).setText( "EMAIL" );
        table.getRow( 0 ).getCell( 2 ).setText( "AGE " );
        //FILL VALUE

        table.getRow( 1 ).getCell( 0).setText( "houssem" );
        table.getRow( 1 ).getCell( 1 ).setText( "test@gmail.com" );
        table.getRow( 1 ).getCell( 2 ).setText( "30" );

*/
        XWPFParagraph paragraph;
        XWPFRun run;
        XWPFTable table;
        XWPFTableRow row;


        //the header
        XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);

        paragraph = header.createParagraph();

        //the body
        paragraph = doc.createParagraph();
        paragraph.setAlignment( ParagraphAlignment.CENTER);
        run = paragraph.createRun();
        run.setText("Patient Report");
        run.setBold(true);
        run.setUnderline( UnderlinePatterns.SINGLE);
        run.setFontSize(18);
        run.setFontFamily("Times New Roman");

       table = doc.createTable(4, 4);
        table.getRows().forEach( x ->  x.getTableCells().forEach(
                t-> t.getParagraphs().forEach( p -> {
                        if(p.getCTP().getPPr() == null) p.getCTP().addNewPPr().addNewKeepLines();
                        p.setKeepNext( true );} )
       ) );

        table.setWidth("100%");
        row = table.getRow(0);
        for(int i = 0  ; i<= 3 ; i++)
            row.getCell( i ).setColor( "9FB4B6" );
        row.getCell(0).setText("CUSTOMER_NAME");
        row.getCell(1).setText("displayName");
        row.getCell(2).setText("CUSTOMER_ID");
        row.getCell(3).setText("displayCustomerID");

        row = table.getRow(1);
        row.getCell(0).setText("x1");
        row.getCell(1).setText("name1");
        row.getCell(2).setText("2022-01-01");
        row.getCell(3).setText("mardi ");

        row = table.getRow(2);

        row.getCell(0).setText("x2");
        row.getCell(1).setText("name2");
        row.getCell(2).setText("2022-01-02");
        row.getCell(3).setText("jeudi ");

        row = table.getRow(3);
        row.getCell(0).setText("x3");
        row.getCell(1).setText("name2");
        row.getCell(2).setText("2022-01-02");
        row.getCell(3).setText("jeudi ");
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

