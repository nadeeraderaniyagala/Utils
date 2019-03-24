# ExcelSheetUtils.class

    private static Workbook book;
    private static TestDataInput testDataInput;

    public static Sheet getTestData(String sheetName) {
        InputStream file;
        try {
            file = new FileInputStream(ConfigFileReader.properties.getProperty("testdata.excel.sheet.path"));
            book = WorkbookFactory.create(file);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        return book.getSheet(sheetName);
    }

    public static String getStringFromCell(Cell cell) {
        return cell != null ? convertNumericCellTypeToString(cell) : null;
    }


    private static String convertNumericCellTypeToString(Cell cell) {
       if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
           cell.setCellType(Cell.CELL_TYPE_STRING);
        }
       return cell.toString();
    }

# ConfigFileReader.class

    public static Properties properties;
    private static FileInputStream inputStream = null;

    static {
        try {
            properties = new Properties();
            inputStream = new FileInputStream("src/test/resources/configs/" + System.getProperty("env") + "/config.properties");
            properties.load(inputStream);
        } catch (
                IOException e) {
            e.printStackTrace();
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

        }
    }

    public static String getReportConfigPath(){
        String reportConfigPath = properties.getProperty("extent.report.configuration.path");
        if(reportConfigPath!= null) return reportConfigPath;
        else throw new RuntimeException("Report Config Path not specified in the Configuration.properties file");
    }
    
    
# Database.util

    public static ResultSet executeQuery(String sqlQuery, Object... args) throws SQLException {
        try (Connection dbConnection =
                     DriverManager.getConnection(properties.getProperty("database.connection.url"),
                             properties.getProperty("database.connection.username"),
                             properties.getProperty("database.connection.password"))) {

            PreparedStatement statement = dbConnection.prepareStatement(sqlQuery);
            for (int i = 1; i < args.length; i++) {
                Object arg = args[i];
                if (arg instanceof String) {
                    String value = (String) arg;
                    statement.setString(i, value);
                } else if (arg instanceof Integer) {
                    Integer value = (Integer) arg;
                    statement.setInt(i, value);
                } else if (arg instanceof Date) {
                    Date value = (Date) arg;
                    statement.setDate(i, value);
                }
            }
            return statement.executeQuery();
        }
    }
