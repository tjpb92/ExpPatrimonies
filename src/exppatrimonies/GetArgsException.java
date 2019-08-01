package exppatrimonies;

/**
 * Classe qui dÃ©finit une exception lancée en cas d'erreur lors de l'analyse des
 * arguments de ligne de commande.
 *
 * @author Thierry Baribaud
 * @version 0.03
 */
public class GetArgsException extends Exception {

    /**
     * Creates a new instance of <code>GetArgsException</code> without detail
     * message.
     */
    public GetArgsException() {
    }

    /**
     * Constructs an instance of <code>GetArgsException</code> with the
     * specified detail message.
     *
     * @param msg the detail message.
     */
    public GetArgsException(String msg) {
        super(msg);
    }
}
