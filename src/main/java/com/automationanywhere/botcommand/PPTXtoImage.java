package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import com.itextpdf.text.Document;
import com.itextpdf.text.Rectangle;
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
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "PPTXtoImage", label = "[[PPTXtoImage.label]]",
        node_label = "[[PPTXtoImage.node_label]]", description = "[[PPTXtoImage.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[PPTXtoImage.return_label]]", return_type = DataType.LIST, return_required = true, return_description = "[[PPTXtoImage.return_description]]")
public class PPTXtoImage {
    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<List<Value>> action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[PPTXtoImage.inputFile.label]]")
            //Force PDF selection
            @FileExtension("pptx")
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

            //Set Optional Export Dir
            @Idx(index = "3", type = TEXT)
            @Pkg(label = "[[PPTXtoImage.outputLocation.label]]", description = "[[PPTXtoImage.outputLocation.description]]")
                    String outputPath) {

        //Internal validation, to disallow empty strings. No null check needed as we have NotEmpty on firstString.
        if ("".equals(inputFile.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        if(!inputFile.toUpperCase().endsWith(".PPTX")){
            throw new BotCommandException("Please select a supported file to continue");
        }
        ListValue<?> result = new ListValue();
        List<Value> resultList = new ArrayList();
        String currentImgFilePath = "";
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

//            //Set full path with file name
//            outputPath = outputPath + fileNameWithoutExt + "." ;

            //Convert to Image
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
                currentImgFilePath = String.format(outputPath + fileNameWithoutExt + "_page%05d.%s", i,outputType);
                FileOutputStream out = new FileOutputStream(currentImgFilePath);
                javax.imageio.ImageIO.write(img, outputType, out);
                resultList.add(new StringValue(currentImgFilePath));
                out.close();
                i++;
            }
            inputStream.close();

        } catch (Exception e) {
            throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
        }
        //Return ListValue.
        result.set(resultList);
        return result;
    }
}
