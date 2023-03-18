# SpringDocConvertTask

请基于mall 框架，实现定时任务管理。定时任务包含，excel doc,ppt 格式文件转换为pdf 格式

## 專案簡介

一個使用Spring boot 的定期任務，可以定期的將指定資料夾內的所有docx檔案轉換成pdf檔案，並且不會重複轉換

## 邏輯

先去列出所有指定資料夾的檔案

如果資料夾中的檔案副檔名為docx就再進一步檢查有沒有同名的pdf檔案

如果有變可以轉換


## Code

```
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
```



## 執行成果

![image](https://user-images.githubusercontent.com/27859973/226114180-0da1d480-1364-47e7-a9a5-d30bacd33ff2.png)

可以看出確實沒有重複轉換。

因為有打印出

System.out.println("Skipped file: " + file.getName() + ", already converted");

## 前置作業

1. clone到本地端後，先去更改application.properties

```
my.inputFolder=C:\\doc
my.outputFolder=C:\\doc
```

換成你想要轉換的指定資料夾和輸出的資料夾路徑

2. 請確保路徑中真的有你的指定資料夾

@Scheduled(fixedDelay = 5000) // 

代表每五秒執行一次，可以自行替換成適當的時間


