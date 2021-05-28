package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.tools.imageio.ImageIOUtil;
import org.fit.pdfdom.PDFDomTree;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.PrintWriter;
import java.io.Writer;
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
        name = "PDFtoImage", label = "[[PDFtoImage.label]]",
        node_label = "[[PDFtoImage.node_label]]", description = "[[PDFtoImage.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[PDFtoImage.return_label]]", return_type = STRING, return_required = true)
public class PDFtoImage {
    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[PDFtoImage.inputFile.label]]")
            //Force PDF selection
            @FileExtension("pdf")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty
                    String inputFile,

            //Select Dropdown for File Conversion
            @Idx(index = "2", type = SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "JPEG", value = "jpeg")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "JPG", value = "jpg")),
                    @Idx.Option(index = "2.3", pkg = @Pkg(label = "GIF", value = "gif")),
                    @Idx.Option(index = "2.4", pkg = @Pkg(label = "TIFF", value = "tiff")),
                    @Idx.Option(index = "2.5", pkg = @Pkg(label = "PNG", value = "png")),
            })
            @NotEmpty
            @Pkg(label = "[[PDFtoImage.outputType.label]]")
                    String outputType,

            //Select Dropdown for File Conversion
            @Idx(index = "3", type = SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Color", value = "color")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Grayscale", value = "grayscale")),
                    @Idx.Option(index = "3.3", pkg = @Pkg(label = "Black and White", value = "blackandwhite"))
            })
            @NotEmpty
            @Pkg(label = "[[PDFtoImage.colorFormat.label]]", description = "[[PDFtoImage.colorFormat.description]]")
                    String colorFormat,

            //Set Optional Export Dir
            @Idx(index = "4", type = TEXT)
            @Pkg(label = "[[PDFtoImage.outputLocation.label]]", description = "[[PDFtoImage.outputLocation.description]]")
                    String outputPath) {

        //Internal validation, to disallow empty strings. No null check needed as we have NotEmpty on firstString.
        if ("".equals(inputFile.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        if(!inputFile.toUpperCase().endsWith(".PDF")){
            throw new BotCommandException("Please select a PDF to continue");
        }

        //Create return value
        String firstPath = "";

        //Business logic
        try{
            //Get file name to add to custom path
            Path path = Paths.get(inputFile);
            Path fileName = path.getFileName();
            String fileNameWithoutExt = fileName.toString().replaceAll("(?).pdf","");

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

            //Convert PDF to Image
            PDDocument pdf = PDDocument.load(new File(inputFile));
            PDFRenderer pdfRenderer = new PDFRenderer(pdf);
            for (int page=0; page < pdf.getNumberOfPages(); ++page){
                if(page==0){
                    //Save file path of file to string for return to UI
                    firstPath = String.format(outputPath + fileNameWithoutExt + "-%05d.%s", page+1,outputType);
                }
                BufferedImage bim;
                if(colorFormat.equals("color")){
                    bim = pdfRenderer.renderImageWithDPI(page, 300, ImageType.RGB);
                }else if(colorFormat.equals("grayscale")){
                    bim = pdfRenderer.renderImageWithDPI(page, 300, ImageType.GRAY);
                }else{
                    bim = pdfRenderer.renderImageWithDPI(page, 300, ImageType.BINARY);
                }

                ImageIOUtil.writeImage(bim, String.format(outputPath + fileNameWithoutExt + "-%05d.%s", page+1,outputType), 300);
            }
            pdf.close();
        } catch (Exception e) {
            throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
        }

        //Return StringValue.
        return new StringValue(firstPath);
    }
}
