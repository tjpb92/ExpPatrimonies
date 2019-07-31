package exppatrimonies;

import bkgpi2a.Agency;
import bkgpi2a.Patrimony;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.mongodb.BasicDBObject;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PaperSize;
import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

/**
 * Programmes servant � exporter dans un fichier Excel les patrimoines extraits
 * d'une base de donn�es MongoDb.
 *
 * @author Thierry Baribaud
 * @version 0.02
 */
public class ExpPatrimonies {

    private final static String path = "c:\\temp";

    private final static String filename = "patrimonies.xlsx";

    private final static String HOST = "1.2.3.4";
    private final static int PORT = 27017;

    /**
     * @param args arguments en ligne de commande
     */
    public static void main(String[] args) {

        FileOutputStream out;
        XSSFWorkbook classeur;
        XSSFSheet feuille;
        XSSFRow titre;
        XSSFCell cell;
        XSSFRow ligne;
        XSSFCellStyle cellStyle;
        XSSFCellStyle titleStyle;
        ObjectMapper objectMapper;
        Patrimony patrimony;
        MongoDatabase mongoDatabase;
        MongoClient MyMongoClient;
        CreationHelper createHelper;
        XSSFHyperlink link;
        XSSFCellStyle hlinkStyle;
        XSSFFont hlinkFont;
        List<String> agencyUuids;
        Agency agency;

        objectMapper = new ObjectMapper();

        MyMongoClient = new MongoClient(HOST, PORT);
        mongoDatabase = MyMongoClient.getDatabase("bdd");

        MongoCollection<Document> patrimoniesCollection = mongoDatabase.getCollection("patrimonies");
        System.out.println(patrimoniesCollection.count() + " patrimoines");

        MongoCollection<Document> agenciesCollection = mongoDatabase.getCollection("agencies");
        System.out.println(agenciesCollection.count() + " agencies");

//      Cr�ation d'un classeur Excel
        classeur = new XSSFWorkbook();
        createHelper = classeur.getCreationHelper();
        feuille = classeur.createSheet("Patrimoines");

        // Style de cellule avec bordure noire
        cellStyle = classeur.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

        // Style pour le titre
        titleStyle = (XSSFCellStyle) cellStyle.clone();
        titleStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        titleStyle.setFillPattern(FillPatternType.LESS_DOTS);
//        titleStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());

        // Style pour les liens dans les cellules
        hlinkStyle = (XSSFCellStyle) cellStyle.clone();
        hlinkFont = classeur.createFont();
        hlinkFont.setUnderline(XSSFFont.U_SINGLE);
        hlinkFont.setColor(HSSFColor.BLUE.index);
        hlinkStyle.setFont(hlinkFont);

        // Ligne de titre
        titre = feuille.createRow(0);
        cell = titre.createCell((short) 0);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("R�f�rence");

        cell = titre.createCell((short) 1);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Label");

        cell = titre.createCell((short) 2);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("ID Performance Immo");

        cell = titre.createCell((short) 3);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Agences");

        // Lit les patrimoines class�s par r�f�rences
        MongoCursor<Document> patrimoniesCursor
                = patrimoniesCollection.find().sort(new BasicDBObject("ref", 1)).iterator();
        int n = 0;
        try {
            while (patrimoniesCursor.hasNext()) {
                patrimony = objectMapper.readValue(patrimoniesCursor.next().toJson(), Patrimony.class);
                System.out.println(n
                        + " ref:" + patrimony.getRef()
                        + ", label:" + patrimony.getLabel()
                        + ", uid:" + patrimony.getUid());
                n++;
                ligne = feuille.createRow(n);

                cell = ligne.createCell(0);
                cell.setCellValue(patrimony.getRef());
                cell.setCellStyle(cellStyle);

                cell = ligne.createCell(1);
                cell.setCellValue(patrimony.getLabel());
                cell.setCellStyle(cellStyle);

                cell = ligne.createCell(2);
                cell.setCellValue(patrimony.getUid());
                link = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.URL);
                link.setAddress("https://dashboard.performance-immo.com/patrimonies/" + patrimony.getUid());
                link.setLabel(patrimony.getUid());
                cell.setHyperlink((XSSFHyperlink) link);
                cell.setCellStyle(hlinkStyle);

                cell = ligne.createCell(3);
                agencyUuids = patrimony.getAgencies();
                if (agencyUuids != null) {
                    MongoCursor<Document> agenciesCursor
                            = agenciesCollection.find(new BasicDBObject("uid", agencyUuids.get(0))).iterator();
                    if (agenciesCursor.hasNext()) {
                        agency = objectMapper.readValue(agenciesCursor.next().toJson(), Agency.class);
                        System.out.println("  label:"+agency.getLabel());
                        cell.setCellValue(agency.getLabel());
                    } else {
                        cell.setCellValue(agencyUuids.get(0));
                    }
                } else {
                    cell.setCellValue("Aucune");
                }
                cell.setCellStyle(cellStyle);
            }

            // Ajustement automatique de la largeur des colonnes
            for (int k = 0; k < 4; k++) {
                feuille.autoSizeColumn(k);
            }

            // Format A4 en sortie
            feuille.getPrintSetup().setPaperSize(PaperSize.A4_PAPER);

            // Orientation paysage
            feuille.getPrintSetup().setLandscape(true);

            // Ajustement � une page en largeur
            feuille.setFitToPage(true);
            feuille.getPrintSetup().setFitWidth((short) 1);
            feuille.getPrintSetup().setFitHeight((short) 0);

            // En-t�te et pied de page
            Header header = feuille.getHeader();
            header.setLeft("Liste des patrimoines Extranet Anstel");
            header.setRight("&F");

            Footer footer = feuille.getFooter();
            footer.setLeft("Documentation confidentielle Anstel");
            footer.setCenter("Page &P / &N");
            footer.setRight("&D");

            // Ligne � r�p�ter en haut de page
            feuille.setRepeatingRows(CellRangeAddress.valueOf("1:1"));

        } catch (IOException ex) {
            Logger.getLogger(ExpPatrimonies.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            patrimoniesCursor.close();
        }

        // Enregistrement du classeur dans un fichier
        try {
            out = new FileOutputStream(new File(path + "\\" + filename));
            classeur.write(out);
            out.close();
            System.out.println("Fichier Excel " + filename + " cr�� dans " + path);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExpPatrimonies.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExpPatrimonies.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

}
