package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.fit.pdfdom.PDFDomTree;

import java.io.File;
import java.io.PrintWriter;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.Future;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "PDFtoHTML", label = "[[PDFtoHTML.label]]",
        node_label = "[[PDFtoHTML.node_label]]", description = "[[PDFtoHTML.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[PDFtoHTML.return_label]]", return_type = STRING, return_required = true)
public class PDFtoHTML {
    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[PDFtoHTML.inputFile.label]]")
            //Force PDF selection
            @FileExtension("pdf")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty
                    String inputFile,

            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[PDFtoHTML.outputLocation.label]]", description = "[[PDFtoHTML.outputLocation.description]]")
                    String outputPath,
            //Select Dropdown for File Conversion
            @Idx(index = "3", type = SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Method 1 - Pure HTML", value = "html")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Method 2 - HTML + Images", value = "image"))
            })
            @NotEmpty
            @Pkg(label = "[[PDFtoHTML.Format.label]]", description = "[[PDFtoHTML.Format.description]]")
                    String conversionMethod) {

        //Internal validation, to disallow empty strings. No null check needed as we have NotEmpty on firstString.
        if ("".equals(inputFile.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        if(!inputFile.toUpperCase().endsWith(".PDF")){
            throw new BotCommandException("Please select a PDF to continue");
        }


        //Business logic
        try{
            //Get file name to add to custom path
            Path path = Paths.get(inputFile);
            Path fileName = path.getFileName();
            String fileNameWithoutExt = fileName.toString().replaceAll("(?).pdf","");

            //Check if output path is same as input or custom
            if(outputPath.equals(null) || outputPath.equals("")){
                //Same as input Path
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
            outputPath = outputPath + fileNameWithoutExt + ".html";

            if(conversionMethod.equals("html")) {

                //Convert PDF to HTML
                PDDocument pdf = PDDocument.load(new File(inputFile));
                Writer output = new PrintWriter(outputPath, "utf-8");
                new PDFDomTree().writeText(pdf, output);
                output.close();
            }else{
                //PDF to HTML in documents4j
                File in = new File(inputFile), target = new File(outputPath);
                IConverter converter = LocalConverter.make();
                Future<Boolean> conversion = converter
                        .convert(in).as(DocumentType.PDF)
                        .to(target).as(DocumentType.HTML)
                        .schedule();
                System.out.println("converter finished");
                converter.shutDown();
            }

        } catch (Exception e) {
            throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
        }

        //Return StringValue.
        return new StringValue(outputPath);
    }
}
