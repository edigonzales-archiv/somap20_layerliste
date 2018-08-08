@Grapes(
    [@Grab(group='org.apache.poi', module='poi', version='3.17'),
    @Grab(group='org.apache.poi', module='poi-ooxml', version='3.17')]
)

import groovy.json.*
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//def apiUrl = new URL("https://geo.so.ch/api/v1/search/layersearch?searchtext=Gemeindegrenze")
//def apiUrl = new URL("https://geo.so.ch/api/v1/search/layersearch?searchtext=Agglomerations")
//def apiUrl = new URL("https://geo.so.ch/api/v1/search/layersearch?searchtext=Bienen")
def apiUrl = new URL("https://geo.so.ch/api/v1/search/layersearch?searchtext=ch")
def data = new JsonSlurper().parseText(apiUrl.text)
def publishedLayers = [:]


data.results.each { result ->
    //println "Titel (=Anzeige im Suchfenster): " + result.title
    //println "ID: " + result.id
    //println "layer.name: " + result.layer.name
    //println "layer.sublayers (map): " + result.layer.sublayers
    //println "layer.sublayers.title: " + result.layer.sublayers.title
    //println "layer.sublayers.sublayers (list): " + result.layer.sublayers.sublayers

    //println "layer.name: " + layer.layer.name
    //println "layer.sublayers: " + layer.layer.sublayers
    //println "layer.sublayers.name: " + layer.layer.sublayers.name
    
    
    def subsublayers = result.layer.sublayers.sublayers[0]
    if (subsublayers != null) {
        subsublayers.each { subsublayer ->
            def publishedLayer = [:]
            publishedLayer.layer = subsublayer.title
            publishedLayer.name = subsublayer.name
            publishedLayer.layergroup = null

            // Falls Title von result und subsublayer identisch ist, dann ist 
            // es keine Layergruppe.
            if (!subsublayer.title.equals(result.title)) {
                publishedLayer.layergroup = result.title
            }

            // Falls Layer noch nicht in Map ist, dann wird er in die Map
            // gespeichert. Falls bereits in Map, muss geprüft werden,
            // ob es eine "Keine-Layergruppe"-Layer ist und das bestehende
            // wird ersetzt.
            def myLayer = publishedLayers.get(subsublayer.title)
            if (myLayer == null || myLayer.layergroup == null ) {
                publishedLayers.put(publishedLayer.layer, publishedLayer)
            }
        }
    } else {
        // Singlelayer oder so!?
        def publishedLayer = [:]
        publishedLayer.layer = result.title
        publishedLayer.name = result.layer.sublayers[0].name
        def myLayer = publishedLayers.get(result.title)
        if (myLayer == null || myLayer.layergroup == null ) {
            publishedLayers.put(publishedLayer.layer, publishedLayer)
        }
    }
}

Workbook workbook = new XSSFWorkbook()
Sheet sheet = workbook.createSheet("Liste")

/*
Row headerRow = sheet.createRow(0)
Cell cell1 = headerRow.createCell(0)
cell1.setCellValue("Kartenebene")
Cell cell2 = headerRow.createCell(1)
cell2.setCellValue("Karte")
Cell cell3 = headerRow.createCell(2)
cell3.setCellValue("Kartenname")
*/

//int rowNum = 1
int rowNum = 0
def entries = publishedLayers.entrySet() 
entries.each { entry ->
    Row row = sheet.createRow(rowNum++);
    try {
        println entry.key + "," + entry.value.layergroup + "," + entry.value.name
        row.createCell(0).setCellValue(entry.key)
        row.createCell(1).setCellValue(entry.value.layergroup)
        row.createCell(2).setCellValue(entry.value.name)
    } catch (java.lang.NullPointerException e) {
        println entry.key + "," + " " + "," + entry.value.name
        row.createCell(0).setCellValue(entry.key)
        row.createCell(1).setCellValue("")
        row.createCell(2).setCellValue(entry.value.name)
    }
}

FileOutputStream fileOut = new FileOutputStream("webgis_client_layerliste.xlsx");
workbook.write(fileOut)
fileOut.close()
workbook.close()
