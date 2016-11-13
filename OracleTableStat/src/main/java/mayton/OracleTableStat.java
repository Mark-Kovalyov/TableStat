package mayton;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.annotation.Nonnull;
import java.io.FileOutputStream;
import java.sql.*;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.logging.Logger;

import static java.lang.String.format;

public class OracleTableStat {

    static final int HEAD_ROW   = 1;
    static final int ROW_OFFSET = 2;
    static final int COL_OFFSET = 0;

    static Logger logger = Logger.getLogger("OracleTableStat");

    static Random random = new Random();

    public static void processTables(@Nonnull Collection<String> tables, @Nonnull Connection conn,
                                     @Nonnull Sheet sheet, int colnum,
                                     @Nonnull String tag, @Nonnull String filter) throws SQLException {


        XSSFCellStyle oracngeStyle = (XSSFCellStyle) sheet.getWorkbook().createCellStyle();
        oracngeStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        XSSFCellStyle greenStyle = (XSSFCellStyle) sheet.getWorkbook().createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());

        int n = 0;
        Row headRow = (sheet.getRow(HEAD_ROW) == null) ? sheet.createRow(HEAD_ROW) : sheet.getRow(HEAD_ROW);
        headRow.createCell(0).setCellValue("Table Name");
        for (String table : tables) {
            headRow.createCell(n*2 + 1).setCellValue("Row count");
            headRow.createCell(n*2 + 2).setCellValue("max(scn)");
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery("select count(*) \"CNT\",max(ORA_ROWSCN) \"MAX_SCN\" from \"" + table + "\"");
            rs.next();
            long cnt = rs.getLong("CNT");
            String scn = rs.getString("MAX_SCN");
            logger.info(format("%s : cnt = %d, scn = %s", table, cnt, scn));
            rs.close();
            st.close();
            int rownum = n + ROW_OFFSET;
            Row row = (sheet.getRow(rownum) == null) ? sheet.createRow(rownum) : sheet.getRow(rownum);
            if (colnum == 0) {
                Cell cell = row.createCell(colnum + COL_OFFSET);
                cell.setCellValue(table);
            }
            Cell cellCnt = row.createCell((colnum*2) + 1 + COL_OFFSET);
            /*if (colnum > 0) {
                Cell prevCell = row.getCell(colnum + COL_OFFSET);
                long prevCnt = (long)prevCell.getNumericCellValue();
                long delta = cnt - prevCnt;
                if (delta > 0) {
                    cellCnt.setCellStyle(greenStyle);
                } else if (delta < 0){
                    cellCnt.setCellStyle(oracngeStyle);
                }
            }*/
            cellCnt.setCellValue(cnt);
            Cell cellScn = row.createCell((colnum*2) + 2 + COL_OFFSET);
            cellScn.setCellValue(scn);
            n++;
        }
    }

    public static void mainLoop(@Nonnull List<String> list, @Nonnull Connection conn, @Nonnull Sheet sheet,
                                @Nonnull String filter) throws SQLException {
        int i = 0;
        do {
            String tag = "Tag1";
            // TODO: Modify tag
            processTables(list, conn, sheet, i, tag, filter);
            sheet.autoSizeColumn(i);
            i++;
        } while (i < 3);
    }

    public static void main(String[] args) throws Exception {
        String driver  = "oracle.jdbc.OracleDriver";
        String jdbcUrl = "jdbc:oracle:thin:scott/tiger@127.0.0.1:1521/XE";
        String filter  = " table_name like '%' ";
        String defaulSheetName = "Oracle table stat";
        Class.forName(driver);
        logger.info("Get connection...");

        XSSFWorkbook book = new XSSFWorkbook();

        Sheet sheet = book.createSheet(defaulSheetName);

        Connection conn = DriverManager.getConnection(jdbcUrl);

        PreparedStatement pst = conn.prepareStatement(
                "SELECT table_name FROM user_tables ORDER BY 1"
        );

        List<String> list = new ArrayList<>();

        ResultSet rs = pst.executeQuery();
        while (rs.next()) {
            list.add(rs.getString("TABLE_NAME"));
        }
        rs.close();

        pst.close();

        mainLoop(list, conn, sheet, filter);

        logger.info("Close connection");
        conn.close();


        // Write the output to a file
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss").withZone(ZoneId.systemDefault());
        String instantString = formatter.format(Instant.now());
        String file = "Oracle.TS.stat-" + instantString + ".xlsx";
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();
        book.close();
        logger.info("Wrote: "+file);
        logger.info("End");


    }
}
