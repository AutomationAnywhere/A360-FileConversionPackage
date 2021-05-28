package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfDocument;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.tools.imageio.ImageIOUtil;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "IMAGEtoPDF", label = "[[IMAGEtoPDF.label]]",
        node_label = "[[IMAGEtoPDF.node_label]]", description = "[[IMAGEtoPDF.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[IMAGEtoPDF.return_label]]", return_type = STRING, return_required = true)
public class IMAGEtoPDF {
    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[IMAGEtoPDF.inputFile.label]]")
            //Force PDF selection
            @FileExtension("jpeg,jpg,gif,png,tiff")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty
                    String inputFile,

            //Set Optional Export Dir
            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[IMAGEtoPDF.outputLocation.label]]", description = "[[IMAGEtoPDF.outputLocation.description]]")
                    String outputPath) {

        //Internal validation, to disallow empty strings. No null check needed as we have NotEmpty on firstString.
        if ("".equals(inputFile.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

//        if(!inputFile.toUpperCase().endsWith(".JPEG")||!inputFile.toUpperCase().endsWith(".JPG")||!inputFile.toUpperCase().endsWith(".PNG")||!inputFile.toUpperCase().endsWith(".GIF")||!inputFile.toUpperCase().endsWith(".TIFF")){
//            throw new BotCommandException("Please select a support image file to continue");
//        }

        //Business logic
        try{
            //Get file name to add to custom path
            Path path = Paths.get(inputFile);
            Path fileName = path.getFileName();
            String fileNameWithoutExt = "";
            if (fileName.toString().toUpperCase().endsWith(".JPEG")){
                fileNameWithoutExt = fileName.toString().replaceAll("(?).jpeg","");
            }else if(fileName.toString().toUpperCase().endsWith(".JPG")){
                fileNameWithoutExt = fileName.toString().replaceAll("(?).jpg","");
            }else if(fileName.toString().toUpperCase().endsWith(".PNG")){
                fileNameWithoutExt = fileName.toString().replaceAll("(?).png","");
            }else if(fileName.toString().toUpperCase().endsWith(".GIF")){
                fileNameWithoutExt = fileName.toString().replaceAll("(?).gif","");
            }else if(fileName.toString().toUpperCase().endsWith(".TIFF")){
                fileNameWithoutExt = fileName.toString().replaceAll("(?).tiff","");
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


            //Convert Image to PDF
            Document document = new Document();
            FileOutputStream fos = new FileOutputStream(outputPath);

            //Get dimensions of existing image for creating image size in PDF
            Image image = Image.getInstance(inputFile);
            float origWidth = image.getWidth();
            float origHeight = image.getHeight();
            image.scaleToFit(origWidth,origHeight);
            image.setAbsolutePosition(0,0);

            //Write PDF
            //Set size for new page based on original Image
            Rectangle rectangle = new Rectangle(origWidth,origHeight);
            PdfWriter writer = PdfWriter.getInstance(document, fos);
            writer.open();
            document.open();
            //Set page size before adding new page
            document.setPageSize(rectangle);
            document.newPage();
            document.add(image);
            document.close();
            writer.close();
        } catch (Exception e) {
            throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
        }

        //Return StringValue.
        return new StringValue(outputPath);
    }
}
