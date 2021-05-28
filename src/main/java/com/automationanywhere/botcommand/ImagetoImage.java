package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.AttributeType.SELECT;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "ImagetoImage", label = "[[ImagetoImage.label]]",
        node_label = "[[ImagetoImage.node_label]]", description = "[[ImagetoImage.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[ImagetoImage.return_label]]", return_type = STRING, return_required = true, return_description = "[[ImagetoImage.return_description]]")
public class ImagetoImage {
    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[ImagetoImage.inputFile.label]]")
            //Force PDF selection
            @FileExtension("jpeg,jpg,gif,png,tiff")
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
            @Pkg(label = "[[ImagetoImage.outputType.label]]")
                    String outputType,

            //Select Dropdown for File Conversion
            @Idx(index = "3", type = SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Color", value = "color")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Grayscale", value = "grayscale")),
                    @Idx.Option(index = "3.3", pkg = @Pkg(label = "Black and White", value = "blackandwhite"))
            })
            @NotEmpty
            @Pkg(label = "[[ImagetoImage.colorFormat.label]]", description = "[[ImagetoImage.colorFormat.description]]")
                    String colorFormat,

            //Set Optional Export Dir
            @Idx(index = "4", type = TEXT)
            @Pkg(label = "[[ImagetoImage.outputLocation.label]]", description = "[[ImagetoImage.outputLocation.description]]")
                    String outputPath) {

        //Internal validation, to disallow empty strings. No null check needed as we have NotEmpty on firstString.
        if ("".equals(inputFile.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        //Business logic
        try{
            //Get file name to add to custom path
            Path path = Paths.get(inputFile);
            Path fileName = path.getFileName();
            String fileNameWithoutExt = "";
            if (fileName.toString().toUpperCase().endsWith(".JPEG")){
                fileNameWithoutExt = fileName.toString().replaceAll("\\.jpeg","");
            }else if(fileName.toString().toUpperCase().endsWith(".JPG")){
                fileNameWithoutExt = fileName.toString().replaceAll("\\.jpg","");
            }else if(fileName.toString().toUpperCase().endsWith(".PNG")){
                fileNameWithoutExt = fileName.toString().replaceAll("\\.png","");
            }else if(fileName.toString().toUpperCase().endsWith(".GIF")){
                fileNameWithoutExt = fileName.toString().replaceAll("\\.gif","");
            }else if(fileName.toString().toUpperCase().endsWith(".TIFF")){
                fileNameWithoutExt = fileName.toString().replaceAll("\\.tiff","");
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


            System.out.println("File Name without Extension " + fileNameWithoutExt);

            //TIFF IMAGES COULD BE MUTLIPAGE
            if(fileName.toString().toUpperCase().endsWith(".TIFF")){
                InputStream inputStream = new FileInputStream(new File(inputFile));
                ImageInputStream imageInputStream = ImageIO.createImageInputStream(inputStream);
                Iterator<ImageReader> iterator = ImageIO.getImageReaders(imageInputStream);
                if (iterator == null || !iterator.hasNext()){
                    //SINGLE PAGE TIFF
                    //Set full path with file name
                    outputPath = outputPath + fileNameWithoutExt + "." + outputType;
                    //Read input image
                    BufferedImage inputImage = ImageIO.read(new File(inputFile));
                    //Create output Image
                    BufferedImage outputImage;
                    //Prepare the image before writing with same dimensions
                    if(colorFormat.equals("color")){
                        //Set up the Output Image
                        outputImage = new BufferedImage(inputImage.getWidth(), inputImage.getHeight(), BufferedImage.TYPE_INT_RGB);
                    }else if(colorFormat.equals("grayscale")){
                        outputImage = new BufferedImage(inputImage.getWidth(), inputImage.getHeight(), BufferedImage.TYPE_BYTE_GRAY);
                    }else{
                        outputImage = new BufferedImage(inputImage.getWidth(), inputImage.getHeight(), BufferedImage.TYPE_BYTE_BINARY);
                    }

                    outputImage.createGraphics().drawImage(inputImage,0,0, Color.WHITE,null);
                    //Write the export
                    ImageIO.write(outputImage, outputType,new File(outputPath));
                } else{
                    //MULTI-PAGE TIFF
                    //Process through pages
                    ImageReader reader = iterator.next();
                    reader.setInput(imageInputStream);
                    String firstPath = "";
                    int numPage = reader.getNumImages(true);
                    for (int i = 0; i < numPage; i++) {
                        String finalOutputPath = "";
                        try {
                            BufferedImage inputTiff = reader.read(i);
                            //Create output Image
                            BufferedImage outputImage;
                            if(colorFormat.equals("color")){
                                //Set up the Output Image
                                outputImage = new BufferedImage(inputTiff.getWidth(), inputTiff.getHeight(), BufferedImage.TYPE_INT_RGB);
                            }else if(colorFormat.equals("grayscale")){
                                outputImage = new BufferedImage(inputTiff.getWidth(), inputTiff.getHeight(), BufferedImage.TYPE_BYTE_GRAY);
                            }else{
                                outputImage = new BufferedImage(inputTiff.getWidth(), inputTiff.getHeight(), BufferedImage.TYPE_BYTE_BINARY);
                            }
                            outputImage.createGraphics().drawImage(inputTiff,0,0, Color.WHITE,null);
                            //Write the export
                            finalOutputPath = String.format(outputPath + fileNameWithoutExt + "-%05d.%s", i+1,outputType);
                            ImageIO.write(outputImage, outputType,new File(String.valueOf(finalOutputPath)));
                            if (i == 0){
                                firstPath = finalOutputPath;
                            }
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    //for return on MP Tiff
                    outputPath = firstPath;
                }



            } else{
                //All OTHER IMAGES GO HERE
                //Read input image
                BufferedImage inputImage = ImageIO.read(new File(inputFile));
                //Create output Image
                BufferedImage outputImage;
                //Prepare the image before writing with same dimensions
                if(colorFormat.equals("color")){
                    //Set up the Output Image
                    outputImage = new BufferedImage(inputImage.getWidth(), inputImage.getHeight(), BufferedImage.TYPE_INT_RGB);
                }else if(colorFormat.equals("grayscale")){
                    outputImage = new BufferedImage(inputImage.getWidth(), inputImage.getHeight(), BufferedImage.TYPE_BYTE_GRAY);
                }else{
                    outputImage = new BufferedImage(inputImage.getWidth(), inputImage.getHeight(), BufferedImage.TYPE_BYTE_BINARY);
                }

                outputImage.createGraphics().drawImage(inputImage,0,0, Color.WHITE,null);
                //Write the export
                ImageIO.write(outputImage, outputType,new File(outputPath));
            }


//
//
//
//
//            //Convert PDF to Image
//            PDDocument pdf = PDDocument.load(new File(inputFile));
//            PDFRenderer pdfRenderer = new PDFRenderer(pdf);
//            for (int page=0; page < pdf.getNumberOfPages(); ++page){
//                BufferedImage bim;
//                if(colorFormat.equals("color")){
//                    bim = pdfRenderer.renderImageWithDPI(page, 300, ImageType.RGB);
//                }else if(colorFormat.equals("grayscale")){
//                    bim = pdfRenderer.renderImageWithDPI(page, 300, ImageType.GRAY);
//                }else{
//                    bim = pdfRenderer.renderImageWithDPI(page, 300, ImageType.BINARY);
//                }
//
//                ImageIOUtil.writeImage(bim, String.format(outputPath + fileNameWithoutExt + "-%05d.%s", page+1,outputType), 300);
//                //Save file path of file to string for return to UI
//                lastFilePath = String.format(outputPath + fileNameWithoutExt + "-%d.%s", page+1,outputType);
//            }
//            pdf.close();

        } catch (Exception e) {
            throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
        }

        //Return StringValue.
        return new StringValue(outputPath);
    }
}
