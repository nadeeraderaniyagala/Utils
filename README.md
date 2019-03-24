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
