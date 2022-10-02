package Test;

import org.elio.StartApp;

import java.io.IOException;
import java.text.ParseException;
import java.util.HashMap;

/**
 * created by elio on 03/10/2022
 */
public class TestClass {
    public static void main( String[] args ) throws IOException, ParseException {

        StartApp.populateInvoiceMap(new HashMap<>());
    }
}
