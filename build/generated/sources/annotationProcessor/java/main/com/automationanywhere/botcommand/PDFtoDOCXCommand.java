package com.automationanywhere.botcommand;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class PDFtoDOCXCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(PDFtoDOCXCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    PDFtoDOCX command = new PDFtoDOCX();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("inputFile") && parameters.get("inputFile") != null && parameters.get("inputFile").get() != null) {
      convertedParameters.put("inputFile", parameters.get("inputFile").get());
      if(convertedParameters.get("inputFile") !=null && !(convertedParameters.get("inputFile") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","inputFile", "String", parameters.get("inputFile").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("inputFile") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","inputFile"));
    }
    if(convertedParameters.containsKey("inputFile")) {
      String filePath= ((String)convertedParameters.get("inputFile"));
      int lastIndxDot = filePath.lastIndexOf(".");
      if (lastIndxDot == -1 || lastIndxDot >= filePath.length()) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.FileExtension","inputFile","pdf"));
      }
      String fileExtension = filePath.substring(lastIndxDot + 1);
      if(!Arrays.stream("pdf".split(",")).anyMatch(fileExtension::equalsIgnoreCase))  {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.FileExtension","inputFile","pdf"));
      }

    }
    if(parameters.containsKey("outputPath") && parameters.get("outputPath") != null && parameters.get("outputPath").get() != null) {
      convertedParameters.put("outputPath", parameters.get("outputPath").get());
      if(convertedParameters.get("outputPath") !=null && !(convertedParameters.get("outputPath") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","outputPath", "String", parameters.get("outputPath").get().getClass().getSimpleName()));
      }
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("inputFile"),(String)convertedParameters.get("outputPath")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }
}
