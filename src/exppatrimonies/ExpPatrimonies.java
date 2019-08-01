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
import utils.ApplicationProperties;
import utils.DBServer;
import utils.DBServerException;

/**
 * Programmes servant à exporter dans un fichier Excel les patrimoines extraits
 * d'une base de données MongoDb.
 *
 * @author Thierry Baribaud
 * @version 0.03
 */
public class ExpPatrimonies {

    /**
     * mgoDbServerType : prod pour le serveur de production, pre-prod pour le
     * serveur de pré-production. Valeur par défaut : pre-prod.
     */
    private String mgoDbServerType = "pre-prod";

    /**
     * debugMode : fonctionnement du programme en mode debug (true/false).
     * Valeur par défaut : false.
     */
    private static boolean debugMode = false;

    /**
     * testMode : fonctionnement du programme en mode test (true/false). Valeur
     * par défaut : false.
     */
    private static boolean testMode = false;

    /**
     * path : répertoire où sera déposé le fichier des résultats
     */
    private String path = ".";

    /**
     * filename : nom du fichier contenant les résultats
     */
    private String filename = "patrimonies.xlsx";

    /**
     * unum : référence au service d'urgence (identifiant interne)
     */
    private int unum;

    /**
     * clientCompanyUuid : identifiant universel unique du service d'urgence
     */
    private String clientCompanyUuid = null;

    private final static String HOST = "1.2.3.4";
    private final static int PORT = 27017;

    /**
     * Constructeur principal de la classe ExpPatrimonies
     *
     * @param args arguments en ligne de commande
     * @throws GetArgsException en cas d'erreur avec les paramètres en ligne de
     * commande
     * @throws java.io.IOException en cas d'erreur d'entrée/sortie.
     * @throws utils.DBServerException en cas d'erreur avec le serveur de base
     * de données.
     */
    public ExpPatrimonies(String[] args) throws GetArgsException, IOException, DBServerException {
        ApplicationProperties applicationProperties;
        DBServer mgoServer;
        MongoClient mongoClient;
        MongoDatabase mongoDatabase;

        System.out.println("Création d'une instance de ExpPatrimonies ...");

        System.out.println("Analyse des arguments de la ligne de commande ...");
        this.getArgs(args);
        System.out.println("Argument(s) en ligne de commande lus().");

        System.out.println("Lecture des paramètres d'exécution ...");
        applicationProperties = new ApplicationProperties("ExpPatrimonies.prop");
        System.out.println("Paramètres d'exécution lus.");

        System.out.println("Lecture des paramètres du serveur Mongo ...");
        mgoServer = new DBServer(mgoDbServerType, "mgoserver", applicationProperties);
        System.out.println("Paramètres du serveur Mongo lus.");
        if (debugMode) {
            System.out.println(mgoServer);
        }

        if (debugMode) {
            System.out.println(this.toString());
        }

        System.out.println("Ouverture de la connexion au serveur MongoDb : " + mgoServer.getName());
        mongoClient = new MongoClient(mgoServer.getIpAddress(), (int) mgoServer.getPortNumber());

        System.out.println("Connexion à la base de données : " + mgoServer.getDbName());
        mongoDatabase = mongoClient.getDatabase(mgoServer.getDbName());

        System.out.println("Export des données ...");
        exportPatrimoniesToExcel(mongoDatabase);

    }

    /**
     * @param args arguments en ligne de commande
     */
    public static void main(String[] args) {
        ExpPatrimonies expPatrimonies;

        System.out.println("Lancement de ExpPatrimonies ...");
        try {
            expPatrimonies = new ExpPatrimonies(args);
        } catch (GetArgsException | IOException | DBServerException exception) {
            Logger.getLogger(ExpPatrimonies.class.getName()).log(Level.SEVERE, null, exception);
//            Logger.getLogger(ExpPatrimonies.class.getName()).log(Level.INFO, null, exception);
        }

        System.out.println("Fin de ExpPatrimonies.");

    }

    /**
     * Récupère les paramètres en ligne de commande
     *
     * @param args arguments en ligne de commande
     */
    private void getArgs(String[] args) throws GetArgsException {
        int i;
        int n;
        int ip1;
        String currentParam;
        String nextParam;

        n = args.length;
        System.out.println("nargs=" + n);
        for (i = 0; i < n; i++) {
            System.out.println("args[" + i + "]=" + args[i]);
        }
        i = 0;
        while (i < n) {
//            System.out.println("args[" + i + "]=" + Args[i]);
            currentParam = args[i];
            ip1 = i + 1;
            nextParam = (ip1 < n) ? args[ip1] : null;
            switch (currentParam) {
                case "-mgodb":
                    if (nextParam != null) {
                        if (nextParam.equals("pre-prod") || nextParam.equals("prod")) {
                            this.mgoDbServerType = nextParam;
                        } else {
                            throw new GetArgsException("ERREUR : Mauvais serveur Mongo : " + nextParam);
                        }
                        i = ip1;
                    } else {
                        throw new GetArgsException("ERREUR : Serveur Mongo non définie");
                    }
                    break;
                case "-path":
                    if (nextParam != null) {
                        this.path = nextParam;
                        i = ip1;
                    } else {
                        throw new GetArgsException("ERREUR : Répertoire non défini");
                    }
                    break;
                case "-o":
                    if (nextParam != null) {
                        this.filename = nextParam;
                        i = ip1;
                    } else {
                        throw new GetArgsException("ERREUR : Fichier non défini");
                    }
                    break;
                case "-u":
                    if (nextParam != null) {
                        try {
                            this.unum = Integer.parseInt(nextParam);
                            i = ip1;
                        } catch (Exception exception) {
                            throw new GetArgsException("L'identifiant du service d'urgence doit être numérique : " + nextParam);
                        }

                    } else {
                        throw new GetArgsException("ERREUR : Identifiant du service d'urgence non défini");
                    }
                    break;
                case "-clientCompanyUuid":
                    if (nextParam != null) {
                        this.clientCompanyUuid = nextParam;
                        i = ip1;
                    } else {
                        throw new GetArgsException("ERREUR : Identifiant UUID du service d'urgence non défini");
                    }
                    break;
                case "-d":
                    setDebugMode(true);
                    break;
                case "-t":
                    setTestMode(true);
                    break;
                default:
                    usage();
                    throw new GetArgsException("ERREUR : Mauvais paramètre : " + currentParam);
            }
            i++;
        }

        if (unum > 0 && clientCompanyUuid != null) {
            System.out.println("unum:" + unum + ", clientCompanyUuid:" + clientCompanyUuid);
            throw new GetArgsException("ERREUR : Veuillez choisir unum ou uuid");
        }
    }

    /**
     * Affiche le mode d'utilisation du programme.
     */
    public static void usage() {
        System.out.println("Usage : java ExpPatrimonies"
                + " [-mgodb prod|pre-prod]"
                + " [-p path]"
                + " [-o file]"
                + " [-u unum|-clientCompanyUuid uuid]"
                + " [-d] [-t]");
    }

    /**
     * @return mgoDbServerType retourne le type de serveur MongoDb
     */
    private String getMgoDbServerType() {
        return (mgoDbServerType);
    }

    /**
     * @param mgoDbServerType définit le type de serveur MongoDb
     */
    private void setMgoDbServerType(String mgoDbServerType) {
        this.mgoDbServerType = mgoDbServerType;
    }

    /**
     * @return debugMode : retourne le mode de fonctionnement debug.
     */
    public boolean getDebugMode() {
        return (debugMode);
    }

    /**
     * @param debugMode : fonctionnement du programme en mode debug
     * (true/false).
     */
    public void setDebugMode(boolean debugMode) {
        ExpPatrimonies.debugMode = debugMode;
    }

    /**
     * @return testMode : retourne le mode de fonctionnement test.
     */
    public boolean getTestMode() {
        return (testMode);
    }

    /**
     * @param testMode : fonctionnement du programme en mode test (true/false).
     */
    public void setTestMode(boolean testMode) {
        ExpPatrimonies.testMode = testMode;
    }

    /**
     * @return retourne répertoire où sera déposé le fichier des résultats
     */
    public String getPath() {
        return path;
    }

    /**
     * @param path définit répertoire où sera déposé le fichier des résultats
     */
    public void setPath(String path) {
        this.path = path;
    }

    /**
     * @return retourne le nom du fichier contenant les résultats
     */
    public String getFilename() {
        return filename;
    }

    /**
     * @param filename définit le nom du fichier contenant les résultats
     */
    public void setFilename(String filename) {
        this.filename = filename;
    }

    /**
     * @return retourne la référence au service d'urgence (identifiant interne)
     */
    public int getUnum() {
        return unum;
    }

    /**
     * @param unum définit la référence au service d'urgence (identifiant
     * interne)
     */
    public void setUnum(int unum) {
        this.unum = unum;
    }

    /**
     * @return retourne l'identifiant universel unique du service d'urgence
     */
    public String getClientCompanyUuid() {
        return clientCompanyUuid;
    }

    /**
     * @param clientCompanyUuid définit l'identifiant universel unique du
     * service d'urgence
     */
    public void setClientCompanyUuid(String clientCompanyUuid) {
        this.clientCompanyUuid = clientCompanyUuid;
    }

    /**
     * Exporte les données dans le fichier Excel
     */
    private void exportPatrimoniesToExcel(MongoDatabase mongoDatabase) {
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
        CreationHelper createHelper;
        XSSFHyperlink link;
        XSSFCellStyle hlinkStyle;
        XSSFFont hlinkFont;
        List<String> agencyUuids;
        Agency agency;

        objectMapper = new ObjectMapper();

        MongoCollection<Document> patrimoniesCollection = mongoDatabase.getCollection("patrimonies");
        System.out.println(patrimoniesCollection.count() + " patrimoines");

        MongoCollection<Document> agenciesCollection = mongoDatabase.getCollection("agencies");
        System.out.println(agenciesCollection.count() + " agencies");

//      Création d'un classeur Excel
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
        cell.setCellValue("Référence");

        cell = titre.createCell((short) 1);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Label");

        cell = titre.createCell((short) 2);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("ID Performance Immo");

        cell = titre.createCell((short) 3);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Agences");

        // Lit les patrimoines classés par références
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
                        System.out.println("  label:" + agency.getLabel());
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

            // Ajustement à une page en largeur
            feuille.setFitToPage(true);
            feuille.getPrintSetup().setFitWidth((short) 1);
            feuille.getPrintSetup().setFitHeight((short) 0);

            // En-tête et pied de page
            Header header = feuille.getHeader();
            header.setLeft("Liste des patrimoines Extranet Anstel");
            header.setRight("&F");

            Footer footer = feuille.getFooter();
            footer.setLeft("Documentation confidentielle Anstel");
            footer.setCenter("Page &P / &N");
            footer.setRight("&D");

            // Ligne à répéter en haut de page
            feuille.setRepeatingRows(CellRangeAddress.valueOf("1:1"));

        } catch (IOException ex) {
            Logger.getLogger(ExpPatrimonies.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            patrimoniesCursor.close();
        }

        // Enregistrement du classeur dans un fichier
        try {
            out = new FileOutputStream(new File(getPath() + "\\" + getFilename()));
            classeur.write(out);
            out.close();
            System.out.println("Fichier Excel " + filename + " créé dans " + path);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExpPatrimonies.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExpPatrimonies.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    /**
     * Retourne le contenu de ExpPatrimonies
     *
     * @return retourne le contenu de ExpPatrimonies
     */
    @Override
    public String toString() {
        return "ExpPatrimonies:{"
                + "mgoDbServerType:" + mgoDbServerType
                + ", path:" + path
                + ", file:" + filename
                + ", unum:" + unum
                + ", clientCompanyUuid:" + clientCompanyUuid
                + ", debugMode:" + debugMode
                + ", testMode:" + testMode
                + "}";
    }

}
