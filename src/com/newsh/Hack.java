package com.newsh;

import java.io.IOException;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Bogdan on 6/19/2015.
 */
public class Hack {
    static void Hack(String path) {
        URI docxUri = null; // "jar:file:/C:/... .docx"
        String uriName = path.replaceAll("\\\\", "/");
        docxUri = URI.create("jar:file:///"+uriName);
        //System.out.println(docxUri);
        Map<String, String> zipProperties = new HashMap<>();
        zipProperties.put("encoding", "UTF-8");
        try (FileSystem zipFS = FileSystems.newFileSystem(docxUri, zipProperties)) {
            Path documentXmlPath = zipFS.getPath("/word/settings.xml");
            byte[] content = Files.readAllBytes(documentXmlPath);
            String xml = new String(content, StandardCharsets.UTF_8);
            //System.out.println(xml);
            xml = xml.replaceAll("<w:documentProtection.*?/>", "");
            //System.out.println(xml);
            content = xml.getBytes(StandardCharsets.UTF_8);
            Files.delete(documentXmlPath);
            Files.write(documentXmlPath, content);
            zipFS.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
