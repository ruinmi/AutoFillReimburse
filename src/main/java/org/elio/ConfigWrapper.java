package org.elio;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

/**
 * created by elio on 02/10/2022
 */
public class ConfigWrapper {
    public String NAME;
    public String DEPARTMENT;
    public String TEMPLATE_FOLDER_PATH;
    public String TAXI_INVOICE_PATH;
    public int CONTENT_START_ROW;
    public int CONTENT_END_ROW;
    public int CONTENT_START_COL;
    public int HEADER_IGNORE_ROW;
    public int START_DAY_OF_MONTH;
    public int END_DAY_OF_MONTH;
    // FORMULA CELLS
    public int ACCUMULATE_HOR_ROW;
    public int ACCUMULATE_HOR_COL;
    public int ACCUMULATE_VER_ROW;
    public int ACCUMULATE_VER_COL;
    public int ACCUMULATE_SHEET;

    // HEADER
    public int PROJECT_LOCATION_ROW;
    public int PROJECT_NUMBER_ROW;
    // BATCH FILL HEADER
    public String PROJECT_LOCATION;
    public String PROJECT_NUMBER;

    // FEE
    public int TRAIN_ROW;
    public int TAXI_ROW;
    public int FOOD_ROW;
    public int BUSINESS_TRIP_ROW;
    // BATCH FILL FEE
    public int FEE_OF_TRAIN;
    public int FEE_OF_FOOD;
    public int FEE_OF_BUSINESS_TRIP;

    private static final Properties prop = new Properties();
    static {
        try {
            InputStream is = StartApp.class.getResourceAsStream("../../config.properties");
            InputStreamReader isr = null;
            if (is != null) {
                isr = new InputStreamReader(is, StandardCharsets.UTF_8);
            }
            prop.load(isr);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public ConfigWrapper() throws IllegalAccessException {
        populateConfig();
    }
    public void populateConfig() throws IllegalAccessException {
        Field[] fields = this.getClass().getFields();
        for (Field field : fields) {
            String name = field.getName();
            String property = prop.getProperty(name);
            if (property == null) {
                continue;
            }
            if (field.getType().isAssignableFrom(int.class)) {
                int value = Integer.parseInt(property);
                field.set(this, value);
                continue;
            }
            field.set(this, property);
        }
    }
}
