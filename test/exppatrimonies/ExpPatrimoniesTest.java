package exppatrimonies;

import java.io.IOException;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;
import utils.DBServerException;
import utils.GetArgsException;

/**
 * Jeux de test pour tester la classe ExpPatrimonies
 *
 * @author Thierry Baribaud
 * @version 0.04
 */
public class ExpPatrimoniesTest {

    public ExpPatrimoniesTest() {
    }

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() {
    }

    @After
    public void tearDown() {
    }

    /**
     * Test of bad Mongo server type.
     */
    @Test
    public void testBadMgoServerType() {
        String[] args = {"-mgodb", "badmgo"};

        System.out.println("ExpPatrimonies.badMgoServerType");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            assertTrue(ex.getMessage().contains("Mauvais serveur Mongo"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

    /**
     * Test of undefined Mongo server type.
     */
    @Test
    public void testUndefinedMgoServerType() {
        String[] args = {"-mgodb"};

        System.out.println("ExpPatrimonies.undefinedMgoServerType");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            assertTrue(ex.getMessage().contains("Serveur Mongo non défini"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

    /**
     * Test of undefined path.
     */
    @Test
    public void testUndefinedPath() {
        String[] args = {"-path"};

        System.out.println("ExpPatrimonies.undefinedPath");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            assertTrue(ex.getMessage().contains("Répertoire non défini"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

    /**
     * Test of undefined filename.
     */
    @Test
    public void testUndefinedFilename() {
        String[] args = {"-o"};

        System.out.println("ExpPatrimonies.undefinedFilename");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            assertTrue(ex.getMessage().contains("Fichier non défini"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

    /**
     * Test of undefined client identifier.
     */
    @Test
    public void testUndefinedUnum() {
        String[] args = {"-u"};

        System.out.println("ExpPatrimonies.undefinedUnum");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            assertTrue(ex.getMessage().contains("Identifiant du service d'urgence non défini"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

    /**
     * Test of bad client identifier.
     */
    @Test
    public void testUndefinedBadUnum() {
        String[] args = {"-u", "badUnum"};

        System.out.println("ExpPatrimonies.badUnum");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            System.out.println(ex);
            assertTrue(ex.getMessage().contains("identifiant du service d'urgence doit être numérique"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

    /**
     * Test of undefined client unique identifier.
     */
    @Test
    public void testUndefinedUUID() {
        String[] args = {"-clientCompanyUuid"};

        System.out.println("ExpPatrimonies.undefinedUUID");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            assertTrue(ex.getMessage().contains("Identifiant UUID du service d'urgence non défini"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

    /**
     * Test of double client identifiers
     */
    @Test
    public void testUndefinedDoubleIds() {
        String[] args = {"-u", "513", "-clientCompanyUuid", "123456-4567"};

        System.out.println("ExpPatrimonies.doubeIds");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            System.out.println(ex);
            assertTrue(ex.getMessage().contains("Veuillez choisir unum ou uuid"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

    /**
     * Test of bad parameter
     */
    @Test
    public void testBadParameter() {
        String[] args = {"-badParameter"};

        System.out.println("ExpPatrimonies.badParameter");

        ExpPatrimonies instance;
        try {
            instance = new ExpPatrimonies(args);
            fail("Expected GetArgsException");
        } catch (GetArgsException ex) {
            assertTrue(ex.getMessage().contains("Mauvais paramètre"));
        } catch (IOException | DBServerException ex) {
            fail("Expected GetArgsException");
        }
    }

}
