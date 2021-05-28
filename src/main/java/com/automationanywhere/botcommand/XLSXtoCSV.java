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

import java.io.File;
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
        name = "XLSXtoCSV", label = "[[XLSXtoCSV.label]]",
        node_label = "[[XLSXtoCSV.node_label]]", description = "[[XLSXtoCSV.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[XLSXtoCSV.return_label]]", return_type = STRING, return_required = true, return_description = "[[XLSXtoCSV.return_description]]")
public class XLSXtoCSV {
    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[XLSXtoCSV.inputFile.label]]")
            //Force PDF selection
            @FileExtension("xlsx")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty
                    String inputFile,

            //Set Optional Export Dir
            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[XLSXtoCSV.outputLocation.label]]", description = "[[XLSXtoCSV.outputLocation.description]]")
                    String outputPath) {

        //Internal validation, to disallow empty strings. No null check needed as we have NotEmpty on firstString.
        if ("".equals(inputFile.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        if(!inputFile.toUpperCase().endsWith(".XLSX")){
            throw new BotCommandException("Please select a supported file to continue");
        }

        //Business logic
        try{
            //Get file name to add to custom path
            Path path = Paths.get(inputFile);
            Path fileName = path.getFileName();
            String fileNameWithoutExt = "";
            if (fileName.toString().toUpperCase().endsWith(".XLSX")) {
                fileNameWithoutExt = fileName.toString().replaceAll("(?).xlsx", "");
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
            outputPath = outputPath + fileNameWithoutExt + ".csv";

            //Convert to PDF
            File in = new File(inputFile), target = new File(outputPath);
            IConverter converter = LocalConverter.make();
            Future<Boolean> conversion = converter
                    .convert(in).as(DocumentType.XLSX)
                    .to(target).as(DocumentType.CSV)
                    .schedule();

            converter.shutDown();

        } catch (Exception e) {
            throw new BotCommandException("Error occurred during file conversion. Error code: " + e.toString());
        }

        //Return StringValue.
        return new StringValue(outputPath);
    }
}
