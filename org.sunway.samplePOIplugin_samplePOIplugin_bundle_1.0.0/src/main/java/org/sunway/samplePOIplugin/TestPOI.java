/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.sunway.samplePOIplugin;

import java.awt.Color;
import java.io.IOException;
import java.util.Map;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.joget.plugin.base.DefaultPlugin;
import org.joget.plugin.base.PluginProperty;
import org.joget.plugin.base.PluginWebSupport;
import java.io.*;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.sql.DataSource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.joget.apps.app.service.AppUtil;

/**
 *
 * @author mabub
 */
public class TestPOI extends DefaultPlugin implements PluginWebSupport
{

    public String getName()
    {
        return "Test POI Plugin";
    }

    public String getVersion()
    {
        return "1.0.0";
    }

    public String getDescription()
    {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    public PluginProperty[] getPluginProperties()
    {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    public Object execute(Map map)
    {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    public void webService(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException
    {

        String refNo = request.getParameter("refNo");
        String query = "SELECT * FROM app_fd_ta_testingapp Where c_refNo = ?";
        TreeMap<Integer, String[]> excelData = new TreeMap<Integer, String[]>();
        int i = 0;

        try
        {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            Connection con = ds.getConnection();
            PreparedStatement stmt = con.prepareStatement(query);
            stmt.setString(1, refNo);
            ResultSet rSet = stmt.executeQuery();

            while (rSet.next())
            {
                excelData.put(i++, new String[]
                {
                    "Ref No", rSet.getString("c_refNo")
                });
                excelData.put(i++, new String[]
                {
                    "Accounting System No", rSet.getString("c_accNo")
                });
                excelData.put(i++, new String[]
                {
                    "Create By", rSet.getString("c_createBy")
                });
                excelData.put(i++, new String[]
                {
                    "Requesting BU", rSet.getString("c_requestingBU")
                });
                excelData.put(i++, new String[]
                {
                    "Description", rSet.getString("c_description")
                });
                excelData.put(i++, new String[]
                {
                    "Amount", rSet.getString("c_amount")
                });
            }
        } catch (SQLException ex)
        {
            Logger.getLogger(TestPOI.class.getName()).log(Level.SEVERE, null, ex);
        }

        //Create blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create a blank sheet
        XSSFSheet spreadsheet = workbook.createSheet(" Report Info ");
        spreadsheet.addMergedRegion(
                new CellRangeAddress(
                        0, //first row (0-based)
                        0, //last row (0-based)
                        0, //first column (0-based)
                        1 //last column (0-based)
                )
        );
  
        spreadsheet.setColumnWidth(0, 50*256); //width of 100 characters
        spreadsheet.setColumnWidth(1, 50*256); //width of 100 characters
        
        //Create row object
        int rowId = 2;

        Set<Integer> keyid = excelData.keySet();
        
        
        //---------------------------
        XSSFCellStyle boldPlain = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldPlain.setFont(boldFont);
        boldPlain.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        boldPlain.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        boldPlain.setRightBorderColor(IndexedColors.BLACK.getIndex());
        boldPlain.setTopBorderColor(IndexedColors.BLACK.getIndex());
        boldPlain.setBorderLeft(CellStyle.BORDER_THIN);
        boldPlain.setBorderBottom(CellStyle.BORDER_THIN);
        boldPlain.setBorderRight(CellStyle.BORDER_THIN);
        boldPlain.setBorderTop(CellStyle.BORDER_THIN);
   
        XSSFCellStyle boldBrown = workbook.createCellStyle();
        boldBrown.setFillForegroundColor(IndexedColors.BROWN.getIndex());
        boldBrown.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        boldBrown.setFont(boldFont);
        boldBrown.setBorderLeft(CellStyle.BORDER_THIN);
        boldBrown.setBorderBottom(CellStyle.BORDER_THIN);
        boldBrown.setBorderRight(CellStyle.BORDER_THIN);
        boldBrown.setBorderTop(CellStyle.BORDER_THIN);
        boldBrown.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        boldBrown.setRightBorderColor(IndexedColors.BLACK.getIndex());
        boldBrown.setTopBorderColor(IndexedColors.BLACK.getIndex());
        boldBrown.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                
        XSSFCellStyle plainBrown = workbook.createCellStyle();
        plainBrown.setFillForegroundColor(IndexedColors.BROWN.getIndex());
        plainBrown.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        plainBrown.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        plainBrown.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        plainBrown.setRightBorderColor(IndexedColors.BLACK.getIndex());
        plainBrown.setTopBorderColor(IndexedColors.BLACK.getIndex());
        plainBrown.setBorderLeft(CellStyle.BORDER_THIN);
        plainBrown.setBorderBottom(CellStyle.BORDER_THIN);
        plainBrown.setBorderRight(CellStyle.BORDER_THIN);
        plainBrown.setBorderTop(CellStyle.BORDER_THIN);
        
 
        XSSFCellStyle justPlain = workbook.createCellStyle();
        justPlain.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        justPlain.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        justPlain.setRightBorderColor(IndexedColors.BLACK.getIndex());
        justPlain.setTopBorderColor(IndexedColors.BLACK.getIndex());
        justPlain.setBorderLeft(CellStyle.BORDER_THIN);
        justPlain.setBorderBottom(CellStyle.BORDER_THIN);
        justPlain.setBorderRight(CellStyle.BORDER_THIN);
        justPlain.setBorderTop(CellStyle.BORDER_THIN);
        
        XSSFCellStyle headingStyle = workbook.createCellStyle();
        headingStyle.setFillForegroundColor(IndexedColors.BLACK.getIndex());
        headingStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());
        headingStyle.setFont(font);
        headingStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        headingStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headingStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headingStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headingStyle.setBorderLeft(CellStyle.BORDER_THIN);
        headingStyle.setBorderBottom(CellStyle.BORDER_THIN);
        headingStyle.setBorderRight(CellStyle.BORDER_THIN);
        headingStyle.setBorderTop(CellStyle.BORDER_THIN);
        
        
        
        
        XSSFRow row;
        Cell cell;
        row = spreadsheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("Excel Reporting for eBilling");
        //spreadsheet.autoSizeColumn(0);
        cell.setCellStyle(headingStyle);
        
        cell = row.createCell(1);
        cell.setCellStyle(headingStyle);
        
        row = spreadsheet.createRow(1);
        cell = row.createCell(0);
        cell.setCellValue("Field");
        //spreadsheet.autoSizeColumn(0);
        cell.setCellStyle(headingStyle);
        cell = row.createCell(1);
        cell.setCellValue("Description");
        //spreadsheet.autoSizeColumn(1);
        cell.setCellStyle(headingStyle);        
        
        //---------------------------
        
        for (int key : keyid)
        {
            row = spreadsheet.createRow(rowId++);
            String[] rowData = excelData.get(key);
            int cellId = 0;
            if ((rowId - 1) % 2 == 0)
            {
                for (String cellData : rowData)
                {
                    cell = row.createCell(cellId++);
                    cell.setCellValue(cellData);
                    //spreadsheet.autoSizeColumn(cellId - 1);
                    if ((cellId - 1) == 0)
                    {
                        cell.setCellStyle(boldPlain);
                    }
                    else
                    {
                        cell.setCellStyle(justPlain);

                    }
                }
            } 
            else
            {
                for (String cellData : rowData)
                {
                    cell = row.createCell(cellId++);
                    cell.setCellValue(cellData);
                    //spreadsheet.autoSizeColumn(cellId - 1);
                    if ((cellId - 1) == 0)
                    {
                        cell.setCellStyle(boldBrown);
                    }
                    else
                    {
                        cell.setCellStyle(plainBrown);
                    }
                }
            }

        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.write(baos);
        response.reset(); 
        // setting some response headers
        response.setHeader("Expires", "0");
        response.setHeader("Cache-Control","must-revalidate, post-check=0, pre-check=0");
        response.setHeader("Pragma", "public");
        // setting the content type
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        // the contentlength
        response.setContentLength(baos.size());
        // write ByteArrayOutputStream to the ServletOutputStream
        OutputStream os = response.getOutputStream();
        baos.writeTo(os);
        os.flush();
        os.close();
        System.out.println("Writesheet.xlsx written successfully");

    }
}
