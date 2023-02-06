package com.exampleexcel.demoexcelread;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

public class XML_Conversion {
  public static void main(String[] args) throws IOException, ParserConfigurationException, TransformerException {
    Workbook workbook = new XSSFWorkbook("/Users/logeshpandij/Downloads/Data1.xlsx");
    Sheet sheet = workbook.getSheetAt(0);

    DocumentBuilderFactory documentFactory = DocumentBuilderFactory.newInstance();
    DocumentBuilder documentBuilder = documentFactory.newDocumentBuilder();
    Document document = documentBuilder.newDocument();
    Element root = document.createElement("root"); //%r
    document.appendChild(root);
    int count=1;
    ArrayList<String> list=new ArrayList<String>();
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
     
      count++;
      Row row = sheet.getRow(i);
      Element rowElement = document.createElement("datafile");

      root.appendChild(rowElement);
      
      for (int j = 0; j < row.getLastCellNum(); j++) {
        
        if(count==2){

          list.add(String.valueOf(row.getCell(j)).replace(".", "").replace(" ", "")+""); 
          // System.out.println(list);
          System.out.println(list.get(j));

         
        }
       else{ 
        Element cellElement = document.createElement(list.get(j));
        cellElement.appendChild(document.createTextNode(row.getCell(j).toString()));

              // System.out.println(row.getCell(j));
        rowElement.appendChild(cellElement );}
       
       
      }
    }

    TransformerFactory transformerFactory = TransformerFactory.newInstance();

    Transformer transformer = transformerFactory.newTransformer();
  

    DOMSource domSource = new DOMSource(document);
    StreamResult streamResult = new StreamResult(new FileOutputStream("demo.xml"));
    transformer.setOutputProperty(javax.xml.transform.OutputKeys.INDENT, "yes");
    // transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", String.valueOf(1));

    transformer.transform(domSource, streamResult);

    System.out.println("XML file created successfully");
    workbook.close();
  }
}
