import com.automationanywhere.botcommand.*;
import com.automationanywhere.botcommand.data.Value;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.IOException;
import java.util.List;

public class TestCases {
    @Test
    public void testPDFtoDOCX(){
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SamplePDF.pdf";
        String outputPath = "src/main/resources/test_files/Output/PDFtoDOCX";

        PDFtoDOCX pdFtoDOCX = new PDFtoDOCX();

        Value<String> outputFile = pdFtoDOCX.action(inputFile,outputPath);
        Assert.assertEquals(outputFile.toString(), "src/main/resources/test_files/Output/PDFtoDOCX/SamplePDF.docx");
    }
    @Test
    public void testXLSXtoCSV(){
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SampleExcel.xlsx";
        //String inputFile = "src/main/resources/test_files/SampleFilesSource/arcade.jpg";
        String outputPath = "src/main/resources/test_files/Output/XLSXtoCSV";

        XLSXtoCSV xlsXtoCSV = new XLSXtoCSV();

        Value<String> outputFile = xlsXtoCSV.action(inputFile,outputPath);
        Assert.assertEquals(outputFile.toString(), "src/main/resources/test_files/Output/XLSXtoCSV/SampleExcel.csv");
    }
    @Test
    public void testPDFtoHTML(){
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SamplePDF.pdf";
        String outputPath = "src/main/resources/test_files/Output/PDFtoHTML/";

        PDFtoHTML pdFtoHTML = new PDFtoHTML();

        Value<String> outputFile = pdFtoHTML.action(inputFile,outputPath,"image");
        Assert.assertEquals(outputFile.toString(), "src/main/resources/test_files/Output/PDFtoHTML/SamplePDF.html");
    }
    @Test
    public void testIMAGEtoPDF(){
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SampleJPG.jpg";
        String outputPath = "src/main/resources/test_files/Output/ImagetoPDF";

        IMAGEtoPDF imageToPDF = new IMAGEtoPDF();

        Value<String> outputFile = imageToPDF.action(inputFile,outputPath);
        Assert.assertEquals(outputFile.toString(), "src/main/resources/test_files/Output/ImagetoPDF/SampleJPG.pdf");
    }
    @Test
    public void testDOCXtoPDF(){
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SampleWordDoc.docx";
        String outputPath = "src/main/resources/test_files/Output/DOCXtoPDF";

        DOCXtoPDF docXtoPDF = new DOCXtoPDF();

        Value<String> outputFile = docXtoPDF.action(inputFile,outputPath);
        Assert.assertEquals(outputFile.toString(), "src/main/resources/test_files/Output/DOCXtoPDF/SampleWordDoc.pdf");
    }
    @Test
    public void testPPTXtoPDF() {
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SamplePowerpoint.pptx";
        String outputPath = "src/main/resources/test_files/Output/PPTXtoPDF";

        PPTXtoPDF pptxToPDF = new PPTXtoPDF();

        Value<String> outputFile = pptxToPDF.action(inputFile,outputPath);
        Assert.assertEquals(outputFile.toString(), "src/main/resources/test_files/Output/PPTXtoPDF/SamplePowerpoint.pdf");
    }
    @Test
    public void testPDFtoPPTX() {
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SamplePDF.pdf";
        String outputPath = "src/main/resources/test_files/Output/PDFtoPPTX";

        PDFtoPPTX pdFtoPPTX = new PDFtoPPTX();

        Value<String> outputFile = pdFtoPPTX.action(inputFile,outputPath);
        Assert.assertEquals(outputFile.toString(), "src/main/resources/test_files/Output/PDFtoPPTX/SamplePDF.pptx");
    }
    @Test
    public void testPPTXtoImage() {
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SamplePowerpoint.pptx";
        String outputPath = "src/main/resources/test_files/Output/PPTXtoImage";

        PPTXtoImage pptXtoImage = new PPTXtoImage();

        Value<List<Value>> outputFile = pptXtoImage.action(inputFile,"jpg",outputPath);
        Assert.assertEquals(outputFile.get(0).toString(), "src/main/resources/test_files/Output/PPTXtoImage/SamplePowerpoint_page00001.jpg");
    }
    @Test
    public void testImagetoImage() {
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SampleMultipageTIFF.tiff";
        String outputPath = "src/main/resources/test_files/Output/ImagetoImage";

        ImagetoImage imagetoImage = new ImagetoImage();

        Value<String> outputFile = imagetoImage.action(inputFile,"tiff","grayscale", outputPath);
        Assert.assertEquals(outputFile.toString(), "src/main/resources/test_files/Output/ImagetoImage/SampleMultipageTIFF-00001.tiff");
    }
    @Test
    public void testPDFtoImage() {
        String inputFile = "src/main/resources/test_files/SampleFilesSource/SamplePDF.pdf";
        String outputPath = "src/main/resources/test_files/Output/PDFtoImage";

        PDFtoImage pdFtoImage = new PDFtoImage();
        ImagetoImage imagetoImage = new ImagetoImage();

        Value<List<Value>> outputFile = pdFtoImage.action(inputFile,"jpg","color", outputPath);
        Assert.assertEquals(outputFile.get(0).toString(), "src/main/resources/test_files/Output/PDFtoImage/SamplePDF-00001.jpg");
    }
}
