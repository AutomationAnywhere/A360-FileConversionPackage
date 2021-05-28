package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.itextpdf.text.Document;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.imgscalr.Scalr;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.TimeUnit;

import static com.automationanywhere.commandsdk.model.AttributeType.FILE;
import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "PDFtoPPTX", label = "[[PDFtoPPTX.label]]",
        node_label = "[[PDFtoPPTX.node_label]]", description = "[[PDFtoPPTX.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[PDFtoPPTX.return_label]]", return_type = STRING, return_required = true, return_description = "[[PDFtoPPTX.return_description]]")
public class PDFtoPPTX {
    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[PDFtoPPTX.inputFile.label]]")
            //Force PDF selection
            @FileExtension("pdf")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty
                    String inputFile,

            //Set Optional Export Dir
            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[PDFtoPPTX.outputLocation.label]]", description = "[[PDFtoPPTX.outputLocation.description]]")
                    String outputPath) {

        //Internal validation, to disallow empty strings. No null check needed as we have NotEmpty on firstString.
        if ("".equals(inputFile.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        if(!inputFile.toUpperCase().endsWith(".PDF")){
            throw new BotCommandException("Please select a supported file to continue");
        }

        //Set temp dir for PPTX to PNG conversion
        String tempPath = "C:\\temp\\fileconversion";
        //Setup Temp Dirs for Conversion
        try{
            if(new File(tempPath).exists()){
                FileUtils.deleteDirectory(new File(tempPath));
            }
            //Create temp-file directories if they dont already exist
            Files.createDirectories(Paths.get(tempPath+"\\images"));
        }catch (Exception e){
            throw new BotCommandException("Error creating temp directories needed for file conversion. Attempt to create C:\\temp\\fileconversion failed. Please check permissions of user account running the bot.");
        }


        //Business logic
        try{
            //Get file name to add to custom path
            Path path = Paths.get(inputFile);
            Path fileName = path.getFileName();
            String fileNameWithoutExt = "";
            if (fileName.toString().toUpperCase().endsWith(".PDF")) {
                fileNameWithoutExt = fileName.toString().replaceAll("(?).pdf", "");
            }

            //Check if output path is same as input or custom
            if(outputPath.equals(null) || outputPath.equals("")){
                //Same as input Path - just remove the file name itself
                outputPath = inputFile.replace(fileName.toString(), "");
            }else{
                //Custom Path
                //Make sure it ends in a slash
                if (!outputPath.endsWith("\\") && outputPath.contains("\\")){
                    outputPath = outputPath + "\\";
                }else if(!outputPath.endsWith("/") && outputPath.contains("/")){
                    outputPath = outputPath + "/";
                }
            }

            //Create file directories if they dont already exist
            Files.createDirectories(Paths.get(outputPath));

            //Set full path with file name
            outputPath = outputPath + fileNameWithoutExt + ".pptx";

            //Convert to PDF
            XMLSlideShow ppt = new XMLSlideShow();
            File sourceFile = new File(inputFile);

            PDDocument document = PDDocument.load(sourceFile);
            PDPageTree list = document.getPages();
            System.out.println("Total files converted -> "+ list.getCount());

            PDFRenderer pdfRenderer = new PDFRenderer(document);
            for(int i=0; i < list.getCount(); i++){
                PDPage page = list.get(i);
                BufferedImage image = pdfRenderer.renderImageWithDPI(i,300, ImageType.RGB);
                Dimension imageMaxSize = new Dimension(1920,1080);
                BufferedImage resizedImage = Scalr.resize(image, Scalr.Method.QUALITY,imageMaxSize.width, imageMaxSize.height);
                File outputFile = new File(tempPath + "\\images\\" + (i+1) + ".png");
                ImageIO.write(resizedImage, "png", outputFile);
                System.out.println("Scaled Image Written Out");

                ppt.setPageSize(new java.awt.Dimension(1920,1080));
                XSLFSlide slide = ppt.createSlide();
                FileInputStream getImage = new FileInputStream(outputFile.getAbsolutePath());
                byte[] pictureData = IOUtils.toByteArray(getImage);

                XSLFPictureData pd = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
                XSLFPictureShape pic = slide.createPicture(pd);
                pic.setAnchor(new java.awt.Rectangle(0,0,resizedImage.getWidth(),resizedImage.getHeight()));
                getImage.close();
            }
            document.close();
            FileOutputStream out = new FileOutputStream(outputPath);
            ppt.write(out);
            out.close();
            ppt.close();
            System.out.println("All things closed");



        } catch (Exception e) {
            throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
        } finally {
            try{
                if(new File(tempPath).exists()){
                    FileUtils.deleteDirectory(new File(tempPath));
                }
            }catch  (Exception e) {
                System.out.println(e.toString());
            }
        }
        //Return StringValue.
        return new StringValue(outputPath);
    }
}
