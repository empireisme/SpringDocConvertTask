package com.mikechen.demo.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

@Component
public class WordToPDFConverter {
	
	@Value("${my.inputFolder}")
	String inputFolder="C:\\doc";
	
	@Value("${my.outputFolder}")
	String outputFolder="C:\\doc";
	
    @Scheduled(fixedDelay = 5000) // 每5秒执行一次
    public void convertWordToPDF() throws Exception {
      
        File folder = new File(inputFolder);
        File[] listOfFiles = folder.listFiles();
        for (File file : listOfFiles) {
            if (file.isFile() && file.getName().endsWith(".docx")) {
               
                File outputFile = new File(outputFolder + "\\" + file.getName().replace(".docx", ".pdf"));
                if (!outputFile.exists()) {
                    InputStream in = new FileInputStream(file);
                	
        	        XWPFDocument document = new XWPFDocument(in);

        	        // Create a PDF document
        	        OutputStream out = new FileOutputStream(outputFile);
        	        Document pdfDocument = new Document();
        	        PdfWriter.getInstance(pdfDocument, out);

        	        // Convert the Word document to PDF
        	        pdfDocument.open();
//        	        pdfDocument.addAuthor("Author Name");
//        	        pdfDocument.addCreator("Creator Name");
//        	        pdfDocument.addSubject("Subject");
//        	        pdfDocument.addTitle("Title");
//        	        pdfDocument.addKeywords("keyword1, keyword2");
//        	        pdfDocument.addCreationDate();
//        	        pdfDocument.addHeader("Header", "Value");

        	        for (XWPFParagraph paragraph : document.getParagraphs()) {
        	            pdfDocument.add(new Paragraph(paragraph.getText()));
        	        }

        	        pdfDocument.close();

        	        System.out.println("PDF file created successfully.");
                    System.out.println("Converted file: " + outputFile.getName());
                } else {
                    System.out.println("Skipped file: " + file.getName() + ", already converted");
                }
            }
        }
    }
    
    public static void main(String[] args) throws Exception {
    	WordToPDFConverter w= new WordToPDFConverter();
    	w.convertWordToPDF();
	}
}
