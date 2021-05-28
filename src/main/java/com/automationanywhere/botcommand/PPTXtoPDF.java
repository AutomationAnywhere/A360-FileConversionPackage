package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import com.itextpdf.text.Document;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.Future;

import static com.automationanywhere.commandsdk.model.AttributeType.FILE;
import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "PPTXtoPDF", label = "[[PPTXtoPDF.label]]",
        node_label = "[[PPTXtoPDF.node_label]]", description = "[[PPTXtoPDF.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[PPTXtoPDF.return_label]]", return_type = STRING, return_required = true, return_description = "[[PPTXtoPDF.return_description]]")
public class PPTXtoPDF {
    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[PPTXtoPDF.inputFile.label]]")
            //Force PDF selection
            @FileExtension("pptx")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty
                    String inputFile,

            //Set Optional Export Dir
            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[PPTXtoPDF.outputLocation.label]]", description = "[[PPTXtoPDF.outputLocation.description]]")
                    String outputPath) {

        //Internal validation, to disallow empty strings. No null check needed as we have NotEmpty on firstString.
        if ("".equals(inputFile.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        if(!inputFile.toUpperCase().endsWith(".PPTX")){
            throw new BotCommandException("Please select a supported file to continue");
        }

        //Set temp dir for PPTX to PNG conversion
        String tempPath = "C:\\temp\\fileconversion";
        //Setup Temp Dirs for Conversion
        try{
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
            if (fileName.toString().toUpperCase().endsWith(".PPTX")) {
                fileNameWithoutExt = fileName.toString().replaceAll("(?).pptx", "");
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
            outputPath = outputPath + fileNameWithoutExt + ".pdf";

            //Convert to PDF
            FileInputStream inputStream = new FileInputStream(inputFile);

            XMLSlideShow ppt = new XMLSlideShow(OPCPackage.open(inputStream));

            inputStream.close();
            Dimension pgsize = ppt.getPageSize();
            float scale = 2;
            int width = (int) (pgsize.width * scale );
            int height = (int) (pgsize.height * scale);
            System.out.println("width:" + width +", height:"+ height);
            int i=1;
            int totalSlides = ppt.getSlides().size();
            System.out.println("totalSlides:" + totalSlides);


            for (XSLFSlide slide : ppt.getSlides()){
                BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0,0,width, height));
                graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
                graphics.setColor(Color.white);
                graphics.clearRect(0,0,width,height);
                graphics.scale(scale,scale);

                slide.draw(graphics);
                FileOutputStream out = new FileOutputStream(tempPath+"\\images\\" + i+".png");
                javax.imageio.ImageIO.write(img, "png", out);
                out.close();
                i++;
            }

            Document document = new Document();
            PdfWriter.getInstance(document, new FileOutputStream(outputPath));
            document.open();
            //create rectangle based on image size for new page
            com.itextpdf.text.Rectangle rectangle = new Rectangle(width,height);
            document.setPageSize(rectangle);
            for(int j = 1; j<=totalSlides; j++){
                com.itextpdf.text.Image slideImage = com.itextpdf.text.Image.getInstance(tempPath+"\\images\\" + j+".png");
                slideImage.scaleToFit(width,height);
                slideImage.setAbsolutePosition(0,0);
                //Add new page and add slide to page
                document.newPage();
                document.add(slideImage);
            }
            document.close();

        } catch (Exception e) {
            throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
        } finally {
            try{
                if(new File(tempPath).exists()){
                    FileUtils.deleteDirectory(new File(tempPath));
                }
            }catch  (Exception e) {
                throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
            }
        }
        //Return StringValue.
        return new StringValue(outputPath);
    }
}
