package ReportTDK;

import java.io.*;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.*;
import jxl.*;
import jxl.biff.formula.FormulaException;
import jxl.format.CellFormat;
import jxl.format.Format;
import jxl.read.biff.BiffException;
import jxl.write.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */

public class Report
{

    protected ResultSet rs;
    private int intDetailX;
    protected int intDetailY;
    protected int intDetailY1;
    private int intColCount;
    private int intTemplateHeight;
    private int intHeaderPos;
    private int intHeaderHeight;
    private int intMaxRow;
    private Vector vtTemplateMerge;
    protected Vector vtGroupList;
    private Vector vtDetailList;
    private Total total;
    private String strTemplateFileName;
    private String strReportPath;
    private Workbook in;
    private WritableSheet sheet;
    protected WritableWorkbook out;
    private Hashtable hstParam;
    private Vector vtColToMerge;
    private OrderCol orderCol;
    protected int intOrder;
    protected int intRow;
    protected Group groups[];
    protected Vector vtTemplate;
    protected Vector vtRS;
    protected Vector vtRow;

    public Report(ResultSet rs, String strTemplateFileName, String strFileName, int intDetailX, int intDetailY, int intTemplateHeight)
        throws IOException, BiffException
    {
        this.rs = null;
        this.intDetailX = 0;
        this.intDetailY = 0;
        intDetailY1 = 0;
        intColCount = 0;
        this.intTemplateHeight = 0;
        intHeaderPos = 0;
        intHeaderHeight = 0;
        intMaxRow = 0;
        vtTemplateMerge = null;
        vtGroupList = null;
        vtDetailList = null;
        total = null;
        this.strTemplateFileName = strTemplateFileName;
        this.strReportPath= strFileName;
        in = null;
        sheet = null;
        out = null;
        hstParam = null;
        vtColToMerge = null;
        orderCol = null;
        intOrder = -1;
        groups = null;
        vtTemplate = null;
        vtRS = null;
        vtRow = null;
        intDetailX--;
        intDetailY--;
        vtRS = new Vector();
        vtRS.add(rs);
        this.rs = rs;
        this.intDetailX = intDetailX;
        this.intDetailY = intDetailY;
        intDetailY1 = intDetailY;
        this.intTemplateHeight = intTemplateHeight;
        vtTemplate = new Vector();
        vtTemplateMerge = new Vector();
        vtGroupList = new Vector();
        vtDetailList = new Vector();
        vtRow = new Vector();
        File fileReport = new File(strFileName);
        File fileTemplate = new File(strTemplateFileName);
        in = Workbook.getWorkbook(fileTemplate);
        out = Workbook.createWorkbook(fileReport, in);
        sheet = out.getSheet(0);
        sheet.getSettings().setAutomaticFormulaCalculation(true);
    }

    public Report(ResultSet rs, String strTemplateFileName, String strFileName)
        throws WriteException, IOException, Exception
    {
        this.rs = null;
        intDetailX = 0;
        intDetailY = 0;
        intDetailY1 = 0;
        intColCount = 0;
        intTemplateHeight = 0;
        intHeaderPos = 0;
        intHeaderHeight = 0;
        intMaxRow = 0;
        vtTemplateMerge = null;
        vtGroupList = null;
        vtDetailList = null;
        total = null;
        this.strTemplateFileName = strTemplateFileName;
        this.strReportPath= strFileName;
        in = null;
        sheet = null;
        out = null;
        hstParam = null;
        vtColToMerge = null;
        orderCol = null;
        intOrder = -1;
        groups = null;
        vtTemplate = null;
        vtRS = null;
        vtRow = null;
        try
        {
            vtRS = new Vector();
            vtRS.add(rs);
            this.rs = rs;
            vtTemplate = new Vector();
            vtTemplateMerge = new Vector();
            hstParam = new Hashtable();
            vtRow = new Vector();
            File fileReport = new File(strFileName);
            File fileTemplate = new File(strTemplateFileName);
            in = Workbook.getWorkbook(fileTemplate);
            out = Workbook.createWorkbook(fileReport, in);
            sheet = out.getSheet(0);
            sheet.getSettings().setAutomaticFormulaCalculation(true);
            parseTemplate();
        }
        catch(Exception ex)
        {
            unInit();
            throw new Exception(ex.getMessage());
        }
    }

    public Report(Vector vtRS, String strTemplateFileName, String strFileName)
        throws WriteException, IOException, Exception
    {
        rs = null;
        intDetailX = 0;
        intDetailY = 0;
        intDetailY1 = 0;
        intColCount = 0;
        intTemplateHeight = 0;
        intHeaderPos = 0;
        intHeaderHeight = 0;
        intMaxRow = 0;
        vtTemplateMerge = null;
        vtGroupList = null;
        vtDetailList = null;
        total = null;
        this.strTemplateFileName = strTemplateFileName;
        this.strReportPath= strFileName;
        in = null;
        sheet = null;
        out = null;
        hstParam = null;
        vtColToMerge = null;
        orderCol = null;
        intOrder = -1;
        groups = null;
        vtTemplate = null;
        this.vtRS = null;
        vtRow = null;
        try
        {
            this.vtRS = vtRS;
            vtTemplate = new Vector();
            vtTemplateMerge = new Vector();
            hstParam = new Hashtable();
            vtRow = new Vector();
            File fileReport = new File(strFileName);
            File fileTemplate = new File(strTemplateFileName);
            in = Workbook.getWorkbook(fileTemplate);
            out = Workbook.createWorkbook(fileReport, in);
            sheet = out.getSheet(0);
            sheet.getSettings().setAutomaticFormulaCalculation(true);
            parseTemplate();
        }
        catch(Exception ex)
        {
            unInit();
            throw new Exception(ex.getMessage());
        }
    }
    //---------------------------------------------------------------------------------------------------------
    //Thiet lap che do bao ve
    //Added Date: 03-09-2008
    //---------------------------------------------------------------------------------------------------------
    public void setProtected(String strPassword)
    {
        sheet.getSettings().setProtected(true);
        sheet.getSettings().setPassword(strPassword);
    }

    public static ResultSet convertVectorToResultSet(Vector vtRows, String strColumnNames[])
    {
        VResultSet vrs = new VResultSet(vtRows);
        vrs.setColumnNames(strColumnNames);
        return vrs;
    }

    private void parseTemplate()
        throws WriteException, IOException, FormulaException, Exception
    {
        int intSheetIndex = 0;
        Template template = null;
        WritableSheet sheet1 = null;
        int intSheetCount = out.getNumberOfSheets();
        if(intSheetCount <= 1)
        {
            throw new Exception("Number of sheets must to be greater than 1");
        }
        sheet1 = out.getSheet(intSheetCount - 1);
        int intRow = 0;
        int intTemplateIndex = 0;
        String strContent = "";
        do
        {
            int intGroupFooterCount = 0;
            rs = (ResultSet)vtRS.get(intTemplateIndex);
            vtGroupList = new Vector();
            vtDetailList = new Vector();
            vtColToMerge = new Vector();
            intDetailX = Integer.parseInt(getCellValue(sheet1, 0, intRow, "0")) - 1;
            intDetailY = Integer.parseInt(getCellValue(sheet1, 1, intRow, "0")) - 1;
            intColCount = Integer.parseInt(getCellValue(sheet1, 2, intRow, "0"));
            if(intColCount <= 0)
            {
                throw new Exception("Column count that is value of sheet" + intSheetCount + ".cell( 3," + (intRow + 1) + ") must to be greater than 0");
            }
            intTemplateHeight = Integer.parseInt(getCellValue(sheet1, 3, intRow, "0"));
            if(intTemplateHeight < 0)
            {
                throw new Exception("The heigh of dynamic section that is value of sheet" + intSheetCount + ".cell( 4," + (intRow + 1) + ") must to be greater or equal  0");
            }
            intHeaderPos = Integer.parseInt(getCellValue(sheet1, 4, intRow, String.valueOf(intDetailY))) - 1;
            intHeaderHeight = Integer.parseInt(getCellValue(sheet1, 5, intRow, "1"));
            intMaxRow = Integer.parseInt(getCellValue(sheet1, 6, intRow, "60000"));
            intRow++;
            do
            {
                strContent = getCellValue(sheet1, 0, intRow, "").trim();
                if(strContent.equals("") || strContent.equalsIgnoreCase("NEXT") || strContent.equalsIgnoreCase("NEXTSHEET"))
                {
                    if(total != null)
                    {
                        total.setGroupFooterCount(intGroupFooterCount);
                    }
                    template = new Template(sheet, intDetailX, intDetailY, intColCount, intTemplateHeight, intHeaderPos, intHeaderHeight, vtGroupList, vtDetailList, total, vtColToMerge, orderCol);
                    template.setSheetIndex(intSheetIndex);
                    vtTemplate.add(template);
                    if(strContent.equalsIgnoreCase("NEXT") || strContent.equalsIgnoreCase("NEXTSHEET")) total= null;
                    if(strContent.equalsIgnoreCase("NEXTSHEET"))
                    {
                        if(++intSheetIndex >= intSheetCount)
                        {
                            throw new Exception("Sheet " + intSheetIndex + " does not exists");
                        }
                        sheet = out.getSheet(intSheetIndex);
                    }
                    break;
                }
                if(strContent.equalsIgnoreCase("Group"))
                {
                    getGroupTemplate(sheet1, intRow);
                }
                if(strContent.equalsIgnoreCase("Detail"))
                {
                    getDetailTemplate(sheet1, intRow);
                }
                if(strContent.equalsIgnoreCase("GroupFooter"))
                {
                    intGroupFooterCount = getGroupFooterTemplate(sheet1, intRow);
                }
                if(strContent.equalsIgnoreCase("Total"))
                {
                    getTotalTemplate(sheet1, intRow);
                }
                if(strContent.equalsIgnoreCase("ColToMerge"))
                {
                    getColToMergeTemplate(sheet1, intRow);
                }
                if(strContent.equalsIgnoreCase("OrderCol"))
                {
                    getOrderColTemplate(sheet1, intRow);
                }
                intRow++;
            } while(true);
            if(strContent.equals(""))
            {
                break;
            }
            intRow++;
            intTemplateIndex++;
        } while(true);
        groups = new Group[vtGroupList.size()];
        for(int inti = 0; inti < vtGroupList.size(); inti++)
        {
            groups[inti] = (Group)vtGroupList.get(inti);
        }

        Range rl1[] = sheet.getMergedCells();
        for(int inti1 = 0; inti1 < rl1.length; inti1++)
        {
            Range r1 = rl1[inti1];
            Cell cell1 = r1.getTopLeft();
            Cell cell2 = r1.getBottomRight();
            if((cell1.getRow() < intDetailY || cell1.getRow() >= intDetailY + intTemplateHeight) && (cell2.getRow() < intDetailY || cell2.getRow() >= intDetailY + intTemplateHeight))
            {
                vtTemplateMerge.add(new Merge(cell1.getColumn(), cell1.getRow(), cell2.getColumn(), cell2.getRow()));
            }
        }

        modifyTemplate();
        sheet = out.getSheet(0);
    }

    public void modifyTemplate()
        throws IOException, WriteException
    {
        HSSFRow row1 = null;
        HSSFCell cell1 = null;
        Cell cell2 = null;
        String strValue1 = "";
        String strValue2 = "";
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(strTemplateFileName));
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        for(int intk = 0; intk < out.getNumberOfSheets(); intk++)
        {
            WritableSheet sheet = out.getSheet(intk);
            int intRowCount = sheet.getRows();
            int intColCount = sheet.getColumns();
            HSSFSheet sheet1 = wb.getSheetAt(intk);
            String strSheetName = wb.getSheetName(intk);
            sheet.setName(strSheetName);
            for(int inti = 0; inti < intRowCount; inti++)
            {
                for(short intj = 0; intj < intColCount; intj++)
                {
                    row1 = sheet1.getRow(inti);
                    if(row1 != null)
                    {
                        cell1 = row1.getCell(intj);
                    }
                    if(cell1 != null && cell1.getCellType() == 1)
                    {
                        strValue1 = cell1.getStringCellValue().toString();
                        parseParam(intk, intj, inti, strValue1);
                    } else
                    {
                        strValue1 = "";
                    }
                    cell2 = sheet.getCell(intj, inti);
                    if(cell2 != null)
                    {
                        strValue2 = cell2.getContents().toString();
                    } else
                    {
                        strValue2 = "";
                    }
                    if(!strValue1.equals("") && !strValue1.equals(strValue2))
                    {
                        Util.addCellToSheet(sheet, cell2, 1, strValue1, intj, inti);
                    }
                }

            }

        }

    }

    private void parseParam(int intSheetIndex, int intCol, int intRow, String strValue)
    {
        if(strValue.length() == 1)
        {
            return;
        }
        int inti = 0;
        do
        {
            if(inti == -1 || inti >= strValue.length())
            {
                break;
            }
            inti = strValue.indexOf("$", inti);
            if(inti == -1)
            {
                break;
            }
            inti++;
            if(strValue.charAt(inti) != '{' || strValue.indexOf("}", inti) == -1)
            {
                int intj = strValue.indexOf(" ", inti + 1);
                if(intj == -1)
                {
                    intj = strValue.indexOf("\n", inti + 1);
                }
                if(intj == -1)
                {
                    intj = strValue.length();
                }
                String strParam = strValue.substring(inti - 1, intj).trim();
                inti = intj;
                if(!strParam.equals("$"))
                {
                    Loc loc = new Loc(intSheetIndex, intCol, intRow);
                    Vector vtLoc;
                    if(hstParam.containsKey(strParam))
                    {
                        vtLoc = (Vector)hstParam.get(strParam);
                    } else
                    {
                        vtLoc = new Vector();
                        hstParam.put(strParam, vtLoc);
                    }
                    vtLoc.add(loc);
                }
            }
        } while(true);
    }

    private void copyTemplateSheetDataToOut(Sheet sheet1, WritableSheet sheet)
        throws WriteException
    {
        String strTContent = "";
        String strContent = "";
        int intRowCount = sheet1.getRows();
        int intColCount = sheet1.getColumns();
        for(int inti = 0; inti < intColCount; inti++)
        {
            for(int intj = 0; intj < intRowCount; intj++)
            {
                Cell tCell = sheet1.getCell(inti, intj);
                if(tCell != null)
                {
                    strTContent = tCell.getContents();
                } else
                {
                    strTContent = "";
                }
                Util.addCellToSheet(sheet, tCell, 1, strTContent, inti, intj);
            }

        }

    }

    public void replaceLastCharacter(String strChar)
        throws WriteException
    {
        int intRowCount = sheet.getRows();
        int intColCount = sheet.getColumns();
        for(int inti = 0; inti < intColCount; inti++)
        {
            for(int intj = 0; intj < intRowCount; intj++)
            {
                Cell cell = sheet.getCell(inti, intj);
                String strContent = cell.getContents();
                if(!strContent.trim().equals("") && strContent.substring(strContent.length() - 1).equals(strChar))
                {
                    strContent = strContent.substring(0, strContent.length() - strChar.length());
                    Util.addCellToSheet(sheet, cell, 1, strContent, inti, intj);
                }
            }

        }

    }

    private String getCellValue(Sheet sheet, int intCol, int intRow, String strDefault)
    {
        String strResult = "";
        Cell cell = sheet.getCell(intCol, intRow);
        if(cell == null)
        {
            return strDefault;
        }
        strResult = cell.getContents();
        if(strResult == null)
        {
            return strDefault;
        }
        if(strResult.equals(""))
        {
            return strDefault;
        } else
        {
            return strResult;
        }
    }

    private String getCellComment(Sheet sheet, int intCol, int intRow, String strDefault)
    {
        String strResult = "";
        Cell cell = sheet.getCell(intCol, intRow);
        if(cell == null)
        {
            return strDefault;
        }
        if(cell.getCellFeatures() == null)
        {
            return strDefault;
        }
        strResult = cell.getCellFeatures().getComment();
        if(strResult == null)
        {
            return strDefault;
        }
        if(strResult.equals(""))
        {
            return strDefault;
        } else
        {
            return strResult;
        }
    }

    private void getGroupTemplate(WritableSheet sheet1, int intRow)
        throws Exception
    {
        String strFromRow = getCellValue(sheet1, 1, intRow, "0");
        if(!Util.isNumeric(strFromRow))
        {
            throw new Exception("The value of cell(2, " + (intRow + 1) + ") must to be numeric");
        }
        int intFromRow = Integer.parseInt(strFromRow);
        if(intFromRow <= 0)
        {
            throw new Exception("The value of cell(2, " + (intRow + 1) + ") must to be greater than 0");
        }
        int intToRow = Integer.parseInt(getCellValue(sheet1, 2, intRow, String.valueOf(intFromRow)));
        for(int inti = intFromRow; inti <= intToRow; inti++)
        {
            Group group = new Group(sheet, rs, inti, intColCount, intDetailX);
            vtGroupList.add(group);
        }

    }

    private void getDetailTemplate(WritableSheet sheet1, int intRow)
        throws Exception
    {
        vtDetailList = new Vector();
        String strFromRow = getCellValue(sheet1, 1, intRow, "0");
        if(!Util.isNumeric(strFromRow))
        {
            throw new Exception("The value of cell(2, " + (intRow + 1) + ") must to be numeric");
        }
        int intFromRow = Integer.parseInt(strFromRow);
        if(intFromRow <= 0)
        {
            throw new Exception("The value of cell(2, " + (intRow + 1) + ") must to be greater than 0");
        }
        int intToRow = Integer.parseInt(getCellValue(sheet1, 2, intRow, String.valueOf(intFromRow)));
        for(int inti = intFromRow; inti <= intToRow; inti++)
        {
            Detail detail = new Detail(sheet, rs, inti, intColCount, intDetailX);
            vtDetailList.add(detail);
        }

    }

    private int getGroupFooterTemplate(WritableSheet sheet1, int intRow)
        throws FormulaException, Exception
    {
        int intGroupFooterCount = 0;
        int inti = 1;
        do
        {
            if(inti <= 0)
            {
                break;
            }
            String strValue = Group.replace(getCellValue(sheet1, inti, intRow, "").trim(), "\n", "");
            if(strValue.equals(""))
            {
                break;
            }
            if(!Util.isNumeric(strValue))
            {
                throw new Exception("The value of Cell(" + (inti + 1) + "," + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + "  must to be numeric");
            }
            int intTotalRow = Integer.parseInt(strValue);
            Total footer = new Total(sheet, rs, intTotalRow, intColCount, intDetailX);
            String strComment = Group.replace(getCellComment(sheet1, inti, intRow, ""), "\n", "");
            if(!Util.isNumeric(strComment))
            {
                throw new Exception("The comment of Cell(" + (inti + 1) + "," + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + "  must to be numeric");
            }
            int intGroupLevel = Integer.parseInt(strComment);
            Group group = getGroupByID(footer.getGroupID());
            if(group == null)
            {
                group = new Group(footer.getGroupID());
                vtGroupList.add(intGroupLevel - 1, group);
            } else
            {
                footer.setTemplateRow1(group.getTemplateRow() + 1);
            }
            group.setGroupFooter(footer);
            intGroupFooterCount++;
            inti++;
        } while(true);
        return intGroupFooterCount;
    }

    private Group getGroupByID(String strGroupID)
    {
        Group group1 = null;
        int inti = 0;
        do
        {
            if(inti >= vtGroupList.size())
            {
                break;
            }
            Group group = (Group)vtGroupList.get(inti);
            if(group.getGroupID().equalsIgnoreCase(strGroupID))
            {
                group1 = group;
                break;
            }
            inti++;
        } while(true);
        return group1;
    }

    private void getTotalTemplate(WritableSheet sheet1, int intRow)
        throws SQLException, FormulaException, Exception
    {
        String strValue = Group.replace(getCellValue(sheet1, 1, intRow, ""), "\n", "");
        if(strValue.equals(""))
        {
            return;
        }
        if(!Util.isNumeric(strValue))
        {
            throw new Exception("The value of Cell(2, " + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + " must to be numeric");
        } else
        {
            int intTotalRow = Integer.parseInt(strValue);
            total = new Total(sheet, rs, intTotalRow, intColCount, intDetailX);
            return;
        }
    }

    private void getParameterTemplate(WritableSheet sheet1, int intRow)
        throws Exception
    {
        int intCol = 1;
        do
        {
            String strContent = getCellValue(sheet1, intCol, intRow, "");
            if(!strContent.equals(""))
            {
                String strLoc = Group.replace(getCellComment(sheet1, intCol, intRow, ""), "\n", "");
                if(strLoc.equals(""))
                {
                    throw new Exception("Comment of Cell(" + (intCol + 1) + ", " + (intRow + 1) + ") that is location of parameter, must to be fill");
                }
                hstParam.put(strContent, strLoc);
                intCol++;
            } else
            {
                return;
            }
        } while(true);
    }

    private void getColToMergeTemplate(WritableSheet sheet1, int intRow)
        throws Exception
    {
        vtColToMerge = new Vector();
        int intCol = 1;
        String strComment = "";
        do
        {
            Cell cell = sheet1.getCell(intCol, intRow);
            String strContent = cell.getContents().trim();
            if(!strContent.equals(""))
            {
                if(!Util.isNumeric(strContent))
                {
                    throw new Exception("The value of Cell(" + (intCol + 1) + ", " + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + " that is column needed to be merged, must to be numeric");
                }
                if(cell.getCellFeatures() == null)
                {
                    strComment = strContent;
                } else
                {
                    strComment = Group.replace(cell.getCellFeatures().getComment().trim(), "\n", "");
                }
                if(!Util.isNumeric(strComment))
                {
                    throw new Exception("The comment of Cell(" + (intCol + 1) + ", " + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + " that is column needed to be compared, must to be numeric");
                }
                int intColToMerge = Integer.parseInt(strContent) - 1;
                int intColToCompare = Integer.parseInt(strComment) - 1;
                ColToMerge coltomerge = new ColToMerge(intColToMerge, intColToCompare);
                vtColToMerge.add(coltomerge);
                intCol++;
            } else
            {
                return;
            }
        } while(true);
    }

    private void getOrderColTemplate(WritableSheet sheet1, int intRow)
        throws Exception
    {
        int intStart = 1;
        int intStep = 1;
        int intIsReset = 0;
        int intColToMerge = 1;
        String strComment = "";
        Cell cell = sheet1.getCell(1, intRow);
        String strContent = cell.getContents().trim();
        if(strContent.equals(""))
        {
            return;
        }
        if(!Util.isNumeric(strContent))
        {
            throw new Exception("The value of Cell(2, " + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + " must to be numeric");
        }
        int intOrderCol = Integer.parseInt(strContent) - 1;
        cell = sheet1.getCell(2, intRow);
        strContent = cell.getContents().trim();
        if(!Util.isNumeric(strContent))
        {
            throw new Exception("The value of Cell(3, " + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + " must to be numeric");
        }
        if(!strContent.equals(""))
        {
            intStart = Integer.parseInt(strContent);
        }
        cell = sheet1.getCell(3, intRow);
        strContent = cell.getContents().trim();
        if(!Util.isNumeric(strContent))
        {
            throw new Exception("The value of Cell(4, " + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + " must to be numeric");
        }
        if(!strContent.equals(""))
        {
            intStep = Integer.parseInt(strContent);
        }
        cell = sheet1.getCell(4, intRow);
        strContent = cell.getContents().trim();
        if(!Util.isNumeric(strContent))
        {
            throw new Exception("The value of Cell(5, " + (intRow + 1) + ") in sheet " + out.getNumberOfSheets() + " must to be numeric");
        }
        if(!strContent.equals(""))
        {
            intIsReset = Integer.parseInt(strContent);
        }
        intColToMerge = intOrderCol;
        cell = sheet1.getCell(5, intRow);
        strContent = getCellValue(sheet1, 5, intRow, "1");
        if(!strContent.equals(""))
        {
            intColToMerge = Integer.parseInt(strContent) - 1;
        }
        orderCol = new OrderCol(intOrderCol, intStart, intStep, intIsReset, intColToMerge);
    }

    public void addGroup(String strGroupFieldName, int intTemplateRow, int intColCount)
        throws FormulaException, Exception
    {
        Group group = new Group(sheet, rs, strGroupFieldName, intTemplateRow, intColCount, intDetailX);
        vtGroupList.add(group);
    }

    public void addDetail(int intTemplateRow, int intColCount)
        throws FormulaException, SQLException, Exception
    {
        Detail detail = new Detail(sheet, rs, intTemplateRow, intColCount, intDetailX);
        vtDetailList.add(detail);
    }

    public void createTotal(int intTemplateRow, int intColCount)
        throws FormulaException, SQLException, Exception
    {
        total = new Total(sheet, rs, intTemplateRow, intColCount, intDetailX);
    }

    public boolean setParameter(String strLoc, String strName, String strValue)
        throws WriteException
    {
        Cell cell1 = sheet.getCell(strLoc);
        if(cell1 == null)
        {
            return false;
        }
        String strCellContent = cell1.getContents();
        String strFormat = cell1.getCellFormat().getFormat().getFormatString();
        int intIndex = strFormat.indexOf("#");
        int intCellType;
        if(intIndex == -1)
        {
            intCellType = 1;
        } else
        {
            intCellType = 2;
        }
        strCellContent = Group.replace(strCellContent, strName.trim(), strValue);
        Util.addCellToSheet(sheet, cell1, intCellType, strCellContent, cell1.getColumn(), cell1.getRow());
        return true;
    }

    public boolean setParameter(String strName, String strValue)
        throws WriteException
    {
        if(hstParam == null)
        {
            return false;
        }
        if(!hstParam.containsKey(strName))
        {
            return false;
        }
        Vector vtLoc = (Vector)hstParam.get(strName);
        for(int inti = 0; inti < vtLoc.size(); inti++)
        {
            Loc loc = (Loc)vtLoc.get(inti);
            int intSheetIndex = loc.intSheetIndex;
            int intCol = loc.intCol;
            int intRow = loc.intRow;
            WritableSheet sheet = out.getSheet(intSheetIndex);
            Cell cell1 = sheet.getCell(intCol, intRow);
            if(cell1 == null)
            {
                continue;
            }
            String strCellContent = cell1.getContents();
            String strFormat = cell1.getCellFormat().getFormat().getFormatString();
            int intIndex = strFormat.indexOf("#");
            int intCellType;
            if(intIndex == -1)
            {
                intCellType = 1;
            } else
            {
                intCellType = 2;
            }
            strCellContent = Group.replace(strCellContent, strName, strValue);
            Util.addCellToSheet(sheet, cell1, intCellType, strCellContent, intCol, intRow);
        }

        return true;
    }
    public boolean setParameter(int intSheetIndex, String strName, String strValue)
        throws WriteException, Exception {
        intSheetIndex--;

        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");

        if(hstParam == null)
        {
            return false;
        }
        if(!hstParam.containsKey(strName))
        {
            return false;
        }
        Vector vtLoc = (Vector)hstParam.get(strName);
        for(int inti = 0; inti < vtLoc.size(); inti++)
        {
            Loc loc = (Loc)vtLoc.get(inti);
            int intCol = loc.intCol;
            int intRow = loc.intRow;
            WritableSheet sheet = out.getSheet(intSheetIndex);
            Cell cell1 = sheet.getCell(intCol, intRow);
            if(cell1 == null)
            {
                continue;
            }
            String strCellContent = cell1.getContents();
            if (strCellContent.trim().equals("")) continue;
            String strFormat = cell1.getCellFormat().getFormat().getFormatString();
            int intIndex = strFormat.indexOf("#");
            int intCellType;
            if(intIndex == -1)
            {
                intCellType = 1;
            } else
            {
                intCellType = 2;
            }
            strCellContent = Group.replace(strCellContent, strName, strValue);
            Util.addCellToSheet(sheet, cell1, intCellType, strCellContent, intCol, intRow);
        }

        return true;
    }
    public void setParameter(Hashtable hstParam)
        throws WriteException
    {
        String strName = "";
        String strValue = "";
        for(Enumeration enum1 = hstParam.keys(); enum1.hasMoreElements(); setParameter(strName, strValue))
        {
            strName = (String)enum1.nextElement();
            strValue = hstParam.get(strName).toString();

            setParameter(strName, strValue);
        }
    }
    public void setParameter(int intSheetIndex, Hashtable hstParam) throws WriteException, Exception {
        String strName = "";
        String strValue = "";
        for(Enumeration enum1 = hstParam.keys(); enum1.hasMoreElements(); setParameter(strName, strValue))
        {
            strName = (String)enum1.nextElement();
            strValue = hstParam.get(strName).toString();

            setParameter(intSheetIndex, strName, strValue);
        }
    }
    public void addColumn(int intCol, int intTemplateCol, int intWidth, Hashtable hstTemplate)
        throws WriteException, SQLException, FormulaException, Exception
    {
        Template template = (Template)vtTemplate.get(0);
        intCol--;
        intTemplateCol--;
        String strTemplateValue = "";
        if(hstTemplate == null)
        {
            return;
        }
        intColCount++;
        template.setColCount(intColCount);
        sheet.insertColumn(intCol);
        if(intWidth != 0)
        {
            sheet.setColumnView(intCol, intWidth);
        } else
        {
            sheet.setColumnView(intCol, sheet.getColumnView(intTemplateCol));
        }
        for(int inti = 0; inti < intHeaderHeight; inti++)
        {
            if(hstTemplate.containsKey("HeaderTitle" + (inti + 1)))
            {
                String strHeaderTitle = hstTemplate.get("HeaderTitle" + (inti + 1)).toString();
                Util.addCellToSheet(sheet, sheet.getCell(intTemplateCol, intHeaderPos), 1, strHeaderTitle, intCol, intHeaderPos + inti);
            }
        }

        for(int inti = 0; inti < vtGroupList.size(); inti++)
        {
            if(hstTemplate.containsKey("Group" + (inti + 1)))
            {
                strTemplateValue = hstTemplate.get("Group" + (inti + 1)).toString();
            } else
            {
                strTemplateValue = "";
            }
            Group group = (Group)vtGroupList.get(inti);
            group.prepairFormulas(intCol);
            group.prepairMerges(intCol);
            group.increaseColCount();
            int intTemplateRow = group.getTemplateRow();
            if(strTemplateValue.indexOf("=") == -1)
            {
                Util.addCellToSheet(sheet, sheet.getCell(intTemplateCol, intTemplateRow), 1, strTemplateValue, intCol, intTemplateRow);
            } else
            {
                strTemplateValue = strTemplateValue.substring(1);
                sheet.addCell(new jxl.write.Formula(intCol, intTemplateRow, strTemplateValue, sheet.getCell(intTemplateCol, intTemplateRow).getCellFormat()));
                group.addFormula(intCol, strTemplateValue);
            }
            group.reParseDetailTemplate();
        }

        for(int inti = 0; inti < vtDetailList.size(); inti++)
        {
            if(hstTemplate.containsKey("Detail" + (inti + 1)))
            {
                strTemplateValue = hstTemplate.get("Detail" + (inti + 1)).toString();
            } else
            {
                strTemplateValue = "";
            }
            Detail detail = (Detail)vtDetailList.get(inti);
            detail.prepairFormulas(intCol);
            detail.prepairMerges(intCol);
            detail.increaseColCount();
            if(strTemplateValue.indexOf("=") == -1)
            {
                Util.addCellToSheet(sheet, sheet.getCell(intTemplateCol, detail.getTemplateRow()), 1, strTemplateValue, intCol, detail.getTemplateRow());
            } else
            {
                strTemplateValue = strTemplateValue.substring(1);
                sheet.addCell(new jxl.write.Formula(intCol, detail.getTemplateRow(), strTemplateValue, sheet.getCell(intTemplateCol, detail.getTemplateRow()).getCellFormat()));
                detail.addFormula(intCol, strTemplateValue);
            }
            detail.reParseDetailTemplate();
        }


        if(total != null)
        {
            if(hstTemplate.containsKey("Total"))
            {
                strTemplateValue = hstTemplate.get("Total").toString();
            } else
            {
                strTemplateValue = "";
            }
            total.prepairFormulas(intCol);
            total.prepairMerges(intCol);
            total.increaseColCount();
            int intTemplateRow = total.getTemplateRow();
            if(strTemplateValue.indexOf("=") == -1)
            {
                Util.addCellToSheet(sheet, sheet.getCell(intTemplateCol, intTemplateRow), 1, strTemplateValue, intCol, intTemplateRow);
            } else
            {
                strTemplateValue = strTemplateValue.substring(1);
                sheet.addCell(new jxl.write.Formula(intCol, intTemplateRow, strTemplateValue, sheet.getCell(intTemplateCol, intTemplateRow).getCellFormat()));
                total.addFormula(intCol, strTemplateValue);
            }
            total.reParseTemplate();
        }
        for(int inti = 0; inti < vtTemplateMerge.size(); inti++)
        {
            Merge merge = (Merge)vtTemplateMerge.get(inti);
            merge.prepairMerge(intCol);
        }

        Range rl1[] = sheet.getMergedCells();
        for(int inti1 = 0; inti1 < rl1.length; inti1++)
        {
            Range r1 = rl1[inti1];
            Cell cell1 = r1.getTopLeft();
            Cell cell2 = r1.getBottomRight();
            if((cell1.getRow() < intDetailY || cell1.getRow() >= intDetailY + intTemplateHeight) && (cell2.getRow() < intDetailY || cell2.getRow() >= intDetailY + intTemplateHeight))
            {
                sheet.unmergeCells(r1);
            }
        }

        Group.addMergeToSheet(sheet, vtTemplateMerge);
    }

    public boolean RemoveGroupTemplate(int intIndex)
    {
        int intCount = vtGroupList.size();
        if(intIndex < 0 && intIndex >= intCount)
        {
            return false;
        }
        for(int intj = intIndex + 1; intj < intCount; intj++)
        {
            Group group = (Group)vtGroupList.get(intj);
            group.intCRow--;
        }

        vtGroupList.remove(intIndex);
        if(total != null)
        {
            total.setTemplateRow1(total.getTemplateRow1() - 1);
        }
        groups = new Group[vtGroupList.size()];
        for(int inti = 0; inti < vtGroupList.size(); inti++)
        {
            groups[inti] = (Group)vtGroupList.get(inti);
        }

        return true;
    }

    public boolean RemoveGroupTemplate(String strGroupID)
    {
        int intCount = vtGroupList.size();
        for(int inti = 0; inti < intCount; inti++)
        {
            Group group = (Group)vtGroupList.get(inti);
            String strGroupID1 = group.getGroupID().trim();
            if(strGroupID1.equalsIgnoreCase(strGroupID.trim()))
            {
                group.unInit();
                for(int intj = inti + 1; intj < intCount; intj++)
                {
                    Group group1 = (Group)vtGroupList.get(intj);
                    group1.intCRow--;
                }

                vtGroupList.remove(inti);
                if(total != null)
                {
                    total.setTemplateRow1(total.getTemplateRow1() - 1);
                }
                return true;
            }
        }

        groups = new Group[vtGroupList.size()];
        for(int inti = 0; inti < vtGroupList.size(); inti++)
        {
            groups[inti] = (Group)vtGroupList.get(inti);
        }

        return false;
    }

    protected void deteteTemplateRow()
    {
        for(int inti = (intDetailY + intTemplateHeight) - 1; inti >= intDetailY; inti--)
        {
            sheet.removeRow(inti);
        }

        intTemplateHeight = vtGroupList.size() + vtDetailList.size();
        if(total != null)
        {
            intTemplateHeight++;
        }
    }

    protected int fillGroupData(int intRow)  throws WriteException, SQLException
    {
        int intGroupCount = vtGroupList.size();
        for(int inti = 0; inti < intGroupCount; inti++)
        {
            if(!groups[inti].isGroupChanged())
            {
                continue;
            }
            for(int intj = intGroupCount - 1; intj >= inti; intj--)
            {
                if(groups[intj].fillFooter(intRow))
                {
                    intRow++;
                }
            }

            if(groups[inti].fillGroup(intRow))
            {
                groups[inti].fillFormula(intRow);
                intRow++;
            }
            clearAllPrevGroup(inti + 1);
        }

        return intRow;
    }

    protected void clearAllPrevGroup(int intFromGroup)
    {
        int intGroupCount = vtGroupList.size();
        for(int intj = intFromGroup; intj < intGroupCount; intj++)
        {
            groups[intj].clearPreviousGroupID();
        }

    }

    protected void clearAllCRowGroup()
    {
        int intGroupCount = vtGroupList.size();
        for(int intj = 0; intj < intGroupCount; intj++)
        {
            groups[intj].intCRow = groups[intj].getTemplateRow();
        }

    }

    protected int fillDetailData(int intRow)
        throws WriteException, SQLException
    {
        int intDetailCount = vtDetailList.size();
        for(int inti = 0; inti < intDetailCount; inti++)
        {
            Detail detail = (Detail)vtDetailList.get(inti);
            detail.fillDetail(intRow);
            intRow++;
        }

        return intRow;
    }

    public void fillDataToExcel() throws SQLException, WriteException, IOException
    {
        fillDataToExcel(true);
    }

    public void fillDataToExcel(boolean isUninit)  throws SQLException, WriteException, IOException
    {
        Template template = null;
        int inti = 0;
        do
        {
            if(inti >= vtTemplate.size())
            {
                break;
            }
            template = (Template)vtTemplate.get(inti);
            rs = (ResultSet)vtRS.get(inti);
            intOrder = -1;
            sheet = template.getSheet();
            int intSheetIndex = template.getSheetIndex();
            intDetailX = template.getDetailX();
            intDetailY = template.getDetailY();
            intDetailY1 = intDetailY;
            intColCount = template.getColCount();
            intTemplateHeight = template.getTemplateHeight();
            intHeaderPos = template.getHeaderPos();
            intHeaderHeight = template.getHeaderHeight();
            vtGroupList = template.getGroupList();
            vtDetailList = template.getDetailList();
            total= template.getTotal();
            groups = new Group[vtGroupList.size()];
            for(int intk = 0; intk < vtGroupList.size(); intk++)
            {
                groups[intk] = (Group)vtGroupList.get(intk);
            }

            total = template.getTotal();
            vtColToMerge = template.getColToMerge();
            orderCol = template.getOrderCol();
            deteteTemplateRow();
            if(rs == null)
            {
                break;
            }
            for(intRow = intDetailY; rs.next(); intRow = fillDetailData(intRow))
            {
                intRow = fillGroupData(intRow);
            }

            fillRemainGroups();
            fillOtherData();
            /*if(intRow == intDetailY)
            {
                intRow--;
            }*/
            vtRow.add(String.valueOf(intRow));
            int intRowCount = (intRow - intDetailY);
            for(int intk = inti + 1; intk < vtTemplate.size(); intk++)
            {
                Template template1 = (Template)vtTemplate.get(intk);
                if(intSheetIndex == template1.getSheetIndex())
                {
                    template1.setDetailY((template1.getDetailY() - intTemplateHeight) + intRowCount);
                }
            }

            inti++;
        } while(true);
        if(isUninit)
        {
            unInit();
        }
    }

    public void fillDataManySheetToExcel()
        throws SQLException, WriteException, IOException
    {
        fillDataManySheetToExcel(true);
    }

    public void fillDataManySheetToExcel(boolean isUninit)
        throws SQLException, WriteException, IOException
    {
        int intSheet = 0;
        if(rs == null)
        {
            unInit();
            return;
        }
        String strSheetName = sheet.getName();
        addSheet(strSheetName + " " + (intSheet + 1), intSheet, intDetailY);
        intRow = intDetailY;
        do
        {
            if(!rs.next())
            {
                break;
            }
            intRow = fillGroupData(intRow);
            intRow = fillDetailData(intRow);
            if(intRow % intMaxRow == 0)
            {
                fillRemainGroups(intRow);
                fillOtherData(intRow);
                for(int inti = sheet.getRows() - 1; inti >= intRow; inti--)
                {
                    sheet.removeRow(inti);
                }

                intSheet++;
                addSheet(strSheetName + " " + (intSheet + 1), intSheet, intDetailY);
                for(int inti = intHeaderPos - 1; inti >= 0; inti--)
                {
                    sheet.removeRow(inti);
                }

                intRow = intHeaderHeight;
                intDetailY1 = intHeaderHeight;
            }
        } while(true);
        fillRemainGroups(intRow);
        fillOtherData(intRow);
        out.removeSheet(0);
        if(isUninit)
        {
            unInit();
        }
    }

    protected void addSheet(String strSheetName, int intIndex, int intRow)
        throws WriteException
    {
        out.copySheet(0, strSheetName, intIndex + 1);
        sheet = out.getSheet(intIndex + 1);
        copyTemplateSheetDataToOut(out.getSheet(0), sheet);
        deteteTemplateRow();
        for(int inti = 0; inti < vtDetailList.size(); inti++)
        {
            Detail detail = (Detail)vtDetailList.get(inti);
            detail.setSheet(sheet);
        }

        for(int inti = 0; inti < vtGroupList.size(); inti++)
        {
            Group group = (Group)vtGroupList.get(inti);
            group.setSheet(sheet);
        }

        if(total != null)
        {
            total.setSheet(sheet);
        }
    }

    protected void fillRemainGroups(int intRow)
        throws SQLException, WriteException
    {
        if(intRow != intDetailY)
        {
            for(int intj = vtGroupList.size() - 1; intj >= 0; intj--)
            {
                if(groups[intj].fillFooter(intRow))
                {
                    intRow++;
                }
                if(groups[intj].intPRow <= groups[intj].intCRow)
                {
                    groups[intj].intPRow = groups[intj].intCRow;
                    groups[intj].intCRow = intRow;
                }
                groups[intj].fillFormula(intRow);
            }

        }
    }

    protected void fillRemainGroups()
        throws SQLException, WriteException
    {
        if(intRow != intDetailY)
        {
            for(int intj = vtGroupList.size() - 1; intj >= 0; intj--)
            {
                if(groups[intj].fillFooter(intRow))
                {
                    intRow++;
                }
                if(groups[intj].intPRow <= groups[intj].intCRow)
                {
                    groups[intj].intPRow = groups[intj].intCRow;
                    groups[intj].intCRow = intRow;
                }
                groups[intj].fillFormula(intRow);
            }

        }
    }

    protected void fillOtherData(int intRow)
        throws SQLException, WriteException
    {
        if(total != null && intRow != intDetailY)
        {
            total.fillData(intRow);
        }
        if(vtColToMerge.size() > 0 && intRow != intDetailY && orderCol == null)
        {
            mergeRowsFollowingCol(intRow);
        }
        if(orderCol != null && intRow != intDetailY && vtColToMerge.size() == 0)
        {
            fillOrderToCol(intRow);
        }
        if(vtColToMerge.size() > 0 && intRow != intDetailY && orderCol != null)
        {
            mergeRowsFollowingColAndFillOrderToCol(intRow);
        }
    }

    protected void fillOtherData()
        throws SQLException, WriteException
    {
        if(total != null && intRow != intDetailY)
        {
            total.fillData(intRow);
            intRow ++;
        }
        if(vtColToMerge.size() > 0 && intRow != intDetailY && orderCol == null)
        {
            mergeRowsFollowingCol(intRow);
        }
        if(orderCol != null && intRow != intDetailY && vtColToMerge.size() == 0)
        {
            fillOrderToCol(intRow);
        }
        if(vtColToMerge.size() > 0 && intRow != intDetailY && orderCol != null)
        {
            mergeRowsFollowingColAndFillOrderToCol(intRow);
        }
    }

    private void mergeRowsFollowingCol(int intRow)
        throws WriteException
    {
        ColToMerge coltomerge = null;
        for(int inti = 0; inti < vtColToMerge.size(); inti++)
        {
            coltomerge = (ColToMerge)vtColToMerge.get(inti);
            int intCol = coltomerge.intColToMerge;
            int intColToCompare = coltomerge.intColToCompare;
            String strStartValue = "";
            String strValue = "";
            String strValue1 = "";
            int intStartRow = intDetailY1;
            int intj = 0;
            for(intj = intDetailY1; intj <= intRow; intj++)
            {
                Cell cell1 = sheet.getCell(intColCount, intj);
                strValue1 = cell1.getContents().toString().trim();
                if(!strValue1.equals("1"))
                {
                    if(intj - intStartRow > 1)
                    {
                        setCellsToBlank(intCol, intStartRow + 1, intj - 1);
                        sheet.mergeCells(intCol, intStartRow, intCol, intj - 1);
                    }
                    intStartRow = intj;
                    strStartValue = "";
                    continue;
                }
                Cell cell2 = sheet.getCell(intColToCompare, intj);
                strValue = cell2.getContents().toString().trim();
                if(strStartValue.equals(""))
                {
                    strStartValue = strValue;
                    intStartRow = intj;
                    continue;
                }
                if(strStartValue.equalsIgnoreCase(strValue))
                {
                    continue;
                }
                if(intj - intStartRow > 1)
                {
                    setCellsToBlank(intCol, intStartRow + 1, intj - 1);
                    sheet.mergeCells(intCol, intStartRow, intCol, intj - 1);
                }
                intStartRow = intj;
                strStartValue = strValue;
            }

            if(!strStartValue.equals("") && intRow - intStartRow > 1)
            {
                setCellsToBlank(intCol, intStartRow + 1, intRow);
                sheet.mergeCells(intCol, intStartRow, intCol, intRow);
            }
        }

    }

    private void setCellsToBlank(int intCol, int intFromRow, int intToRow)
        throws WriteException
    {
        for(int intj = intFromRow; intj <= intToRow; intj++)
        {
            Cell cell1 = sheet.getCell(intCol, intj);
            sheet.addCell(new Blank(intCol, intj, cell1.getCellFormat()));
        }

    }

    private void fillOrderToCol(int intRow)
        throws WriteException
    {
        int intOrderCol = orderCol.intOrderCol;
        int intStart = orderCol.intStart;
        int intStep = orderCol.intStep;
        int intIsReset = orderCol.intIsReset;
        String strValue1 = "";
        Cell cell1 = null;
        if(intOrder == -1)
        {
            intOrder = intStart;
        }
        for(int intj = intDetailY1; intj <= intRow; intj++)
        {
            cell1 = sheet.getCell(intColCount, intj);
            strValue1 = cell1.getContents().toString().trim();
            if(!strValue1.equals("1"))
            {
                if(intIsReset != 0)
                {
                    intOrder = intStart;
                }
            } else
            {
                sheet.addCell(new jxl.write.Number(intOrderCol, intj, intOrder, sheet.getCell(intOrderCol, intj).getCellFormat()));
                intOrder += intStep;
            }
        }

    }

    private void mergeRowsFollowingColAndFillOrderToCol(int intRow)
        throws WriteException
    {
        int intOrderCol = orderCol.intOrderCol;
        int intStart = orderCol.intStart;
        int intStep = orderCol.intStep;
        int intIsReset = orderCol.intIsReset;
        int intColToMerge = orderCol.intColToMerge;
        String strValue1 = "";
        Cell cell1 = null;
        if(intOrder == -1)
        {
            intOrder = intStart;
        }
        ColToMerge coltomerge = null;
        for(int inti = 0; inti < vtColToMerge.size(); inti++)
        {
            coltomerge = (ColToMerge)vtColToMerge.get(inti);
            int intCol = coltomerge.intColToMerge;
            int intColToCompare = coltomerge.intColToCompare;
            String strStartValue = "";
            String strValue = "";
            int intStartRow = intDetailY1;
            int intj = 0;
            for(intj = intDetailY1; intj <= intRow; intj++)
            {
                cell1 = sheet.getCell(intColCount, intj);
                strValue1 = cell1.getContents().toString().trim();
                if(!strValue1.equals("1"))
                {
                    if(intIsReset != 0 && (intColToMerge == intOrderCol && inti == 0 || intColToMerge != intOrderCol && intColToMerge == intCol))
                    {
                        intOrder = intStart;
                    }
                    if(intj - intStartRow > 1)
                    {
                        setCellsToBlank(intCol, intStartRow + 1, intj - 1);
                        sheet.mergeCells(intCol, intStartRow, intCol, intj - 1);
                    }
                    intStartRow = intj;
                    strStartValue = "";
                    continue;
                }
                if(intColToMerge == intOrderCol && inti == 0)
                {
                    sheet.addCell(new jxl.write.Number(intOrderCol, intj, intOrder, sheet.getCell(intOrderCol, intj).getCellFormat()));
                    intOrder += intStep;
                }
                Cell cell2 = sheet.getCell(intColToCompare, intj);
                strValue = cell2.getContents().toString().trim();
                if(strStartValue.equals(""))
                {
                    strStartValue = strValue;
                    intStartRow = intj;
                    if(intColToMerge != intOrderCol && intColToMerge == intCol)
                    {
                        sheet.addCell(new jxl.write.Number(intOrderCol, intj, intOrder, sheet.getCell(intOrderCol, intj).getCellFormat()));
                        intOrder += intStep;
                    }
                    continue;
                }
                if(strStartValue.equalsIgnoreCase(strValue))
                {
                    continue;
                }
                if(intColToMerge != intOrderCol && intColToMerge == intCol)
                {
                    sheet.addCell(new jxl.write.Number(intOrderCol, intj, intOrder, sheet.getCell(intOrderCol, intj).getCellFormat()));
                    intOrder += intStep;
                }
                if(intj - intStartRow > 1)
                {
                    setCellsToBlank(intCol, intStartRow + 1, intj - 1);
                    sheet.mergeCells(intCol, intStartRow, intCol, intj - 1);
                }
                intStartRow = intj;
                strStartValue = strValue;
            }

            if(!strStartValue.equals("") && intRow - intStartRow > 1)
            {
                setCellsToBlank(intCol, intStartRow + 1, intRow);
                sheet.mergeCells(intCol, intStartRow, intCol, intRow);
            }
        }

    }

    public int getCurrentRow()
    {
        return intRow;
    }

    public int getCurrentRow(int intTemplateIndex)
    {
        if(intTemplateIndex < 0 || intTemplateIndex >= vtRow.size())
        {
            return 0;
        } else
        {
            return Integer.parseInt(vtRow.get(intTemplateIndex).toString());
        }
    }

    public String getCellValue(int intCol, int intRow) throws Exception {
        return getCellValue(1, intCol, intRow);
    }
    //--------------------------------------------------------------------------------------------
    //Lay gia tri o
    //Added Date: 20-06-2008
    //--------------------------------------------------------------------------------------------
    public String getCellValue(int intSheetIndex, int intCol, int intRow) throws Exception {
        String strValue = "";
        intRow--;
        intCol--;
        intSheetIndex--;
        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
        Sheet sheet= out.getSheet(intSheetIndex);

        strValue = Util.getCellValue(sheet, intCol, intRow);

        return strValue;
    }
    //--------------------------------------------------------------------------------------------
    public String getCellValue(String strLoc) throws Exception {
       return getCellValue(1, strLoc);
    }
    //--------------------------------------------------------------------------------------------
    //Lay gia tri o
    //Added Date: 20-06-2008
    //--------------------------------------------------------------------------------------------
    public String getCellValue(int intSheetIndex, String strLoc) throws Exception {
        String strValue = "";

        intSheetIndex--;
        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
        Sheet sheet= out.getSheet(intSheetIndex);

        Cell cell = sheet.getCell(strLoc);
        if(cell == null)
        {
            return "";
        } else
        {
            int intRow = cell.getRow();
            int intCol = cell.getColumn();
            strValue = Util.getCellValue(sheet, intCol, intRow);
            return strValue;
        }
    }

    public void setValue(int intCol, int intRow, String strValue) throws WriteException, Exception {
        setValue(1, intCol, intRow, strValue);
    }
    //--------------------------------------------------------------------------------------------
    //Thiet lap GT cho cell
    //Added Date: 20-06-2008
    //--------------------------------------------------------------------------------------------
    public void setValue(int intSheetIndex, int intCol, int intRow, String strValue) throws WriteException, Exception {
        intSheetIndex--;
        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
        WritableSheet sheet= out.getSheet(intSheetIndex);

        Cell cell = sheet.getCell(intCol, intRow);
        if(cell == null)
        {
            return;
        } else
        {
            int intCellType = Util.getCellType(cell);
            Util.addCellToSheet(sheet, cell, intCellType, strValue, intCol, intRow);
            return;
        }
    }
    //--------------------------------------------------------------------------------------------
    //Thiet lap GT cho cell
    //Added Date: 20-06-2008
    //--------------------------------------------------------------------------------------------
    public void setValue( String strLoc, String strValue)  throws WriteException, Exception {
        setValue(1, strLoc, strValue) ;
    }
    //--------------------------------------------------------------------------------------------
    //Thiet lap GT cho cell
    //Added Date: 20-06-2008
    //--------------------------------------------------------------------------------------------
    public void setValue(int intSheetIndex, String strLoc, String strValue)  throws WriteException, Exception {
        intSheetIndex--;
        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
        WritableSheet sheet= out.getSheet(intSheetIndex);

        Cell cell = sheet.getCell(strLoc);
        if(cell == null)
        {
            return;
        } else
        {
            int intCellType = Util.getCellType(cell);
            Util.addCellToSheet(sheet, cell, intCellType, strValue, cell.getColumn(), cell.getRow());
            return;
        }
    }
    //--------------------------------------------------------------------------------------------
    //RemoveFormulas
    //Added Date: 04-10-2008
    //--------------------------------------------------------------------------------------------
    public void removeFormulas() throws Exception {
       Util.removeFormula(sheet);
    }
    //--------------------------------------------------------------------------------------------
    //RemoveFormulas
    //Added Date: 04-10-2008
    //--------------------------------------------------------------------------------------------
    public void removeFormulas(int intSheetIndex) throws Exception {
       if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
       if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
       WritableSheet sheet= out.getSheet(intSheetIndex);

       Util.removeFormula(sheet);
    }
    //--------------------------------------------------------------------------------------------
    //Merge cell
    //Added Date: 20-06-2008
    //--------------------------------------------------------------------------------------------
    public void mergeCells(int intCell1, int intRow1, int intCell2, int intRow2) throws WriteException,      Exception {
        mergeCells(1, intCell1, intRow1, intCell2, intRow2);
    }
    //--------------------------------------------------------------------------------------------
    //Merge cell
    //Added Date: 20-06-2008
    //--------------------------------------------------------------------------------------------
    public void mergeCells(int intSheetIndex, int intCell1, int intRow1, int intCell2, int intRow2) throws WriteException,      Exception {
        intSheetIndex--;
        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
        WritableSheet sheet= out.getSheet(intSheetIndex);

        intCell1--;
        intRow1--;
        intCell2--;
        intRow2--;
        sheet.mergeCells(intCell1, intRow1, intCell2, intRow2);
    }

    public void setTotalCellValue(int intCol, String strValue)
    {
        if(total == null)
        {
            return;
        } else
        {
            total.setCellValue(intCol, strValue);
            return;
        }
    }
    //--------------------------------------------------------------------------------------------
    //Thiet lap cell cho total value
    //Added Date: 20-06-2008
    //--------------------------------------------------------------------------------------------
    public void setTotalCellValue(int intTemplateIndex, int intCol, String strValue) throws Exception {
        if (intTemplateIndex <0 ) throw new Exception("Template index must to be equal or greater than 0");
        if (intTemplateIndex >= vtTemplate.size() ) throw new Exception("intTemplateIndex index must to be less than number of sheets");

        Template template= (Template) vtTemplate.get(intTemplateIndex);
        Total total= template.getTotal();

        if(total == null)
        {
            return;
        } else
        {
            total.setCellValue(intCol, strValue);
            return;
        }
    }

    public void setGroupFooterCellValue(int intGroupLevel, int intCol, String strValue)
    {
        if(intGroupLevel > vtGroupList.size())
        {
            return;
        } else
        {
            Total footer = groups[intGroupLevel - 1].getGroupFooter();
            footer.setCellValue(intCol, strValue);
            return;
        }
    }
    /*//----------------------------------------------------------------------------------------------------
    //Added Date: 09-07-2008
    //----------------------------------------------------------------------------------------------------
    public void insertImage(String strImagePath, int intx1, int inty1, int intx2, int inty2, short shtCol1, int intRow1, short shtCol2, int intRow2)
    {
        insertImage(strImagePath, 0, intx1, inty1, intx2, inty2, shtCol1, intRow1, shtCol2, intRow2);
    }
    //----------------------------------------------------------------------------------------------------
    //Added Date: 09-07-2008
    //----------------------------------------------------------------------------------------------------
    public void insertImage(String strImagePath, int intSheetIndex, int intx1, int inty1, int intx2, int inty2, short shtCol1, int intRow1, short shtCol2, int intRow2)
    {
        POIFSFileSystem fs= null;
        HSSFWorkbook wb= null;

        try
        {
           fs = new POIFSFileSystem(new FileInputStream(strReportPath));
           wb = new HSSFWorkbook(fs);
           HSSFSheet sheet= wb.getSheetAt(intSheetIndex);
           HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
           HSSFClientAnchor anchor= new HSSFClientAnchor(intx1, inty1, intx2, inty2, shtCol1, intRow1, shtCol2, intRow2);
           patriarch.createPicture(anchor, loadPicture(strImagePath, wb));

        }catch(Exception ex)
        {
           ex.printStackTrace();
        }
    }
    //----------------------------------------------------------------------------------------------------
    //Added Date: 09-07-2008
    //----------------------------------------------------------------------------------------------------
    private int loadPicture( String strImagePath, HSSFWorkbook wb ) throws IOException
    {
        int intPictureIndex;
        int intPictureType= 0;

        if (strImagePath.toUpperCase().endsWith(".PNG")) intPictureType= HSSFWorkbook.PICTURE_TYPE_PNG;
        if (strImagePath.toUpperCase().endsWith(".JPG")) intPictureType= HSSFWorkbook.PICTURE_TYPE_JPEG;
        if (strImagePath.toUpperCase().endsWith(".BMP")) intPictureType= HSSFWorkbook.PICTURE_TYPE_DIB;

        FileInputStream fis = null;
        ByteArrayOutputStream bos = null;
        try
        {
            fis = new FileInputStream(strImagePath);
            bos = new ByteArrayOutputStream( );
            int intc;
            while ( (intc = fis.read()) != -1)
                bos.write( intc );

            intPictureIndex = wb.addPicture( bos.toByteArray(), intPictureType);
        }
        finally
        {
            if (fis != null) fis.close();
            if (bos != null) bos.close();
        }
        return intPictureIndex;
    }*/
    //----------------------------------------------------------------------------------------------------
    //Them anh toi
    //Added Date: 20-06-2008
    //----------------------------------------------------------------------------------------------------
    public void insertImage(int intSheetIndex, File file, int intCol, int intRow, int intWidth, int intHeight) throws Exception {
        intSheetIndex--;
        intCol--;
        intRow--;
        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
        WritableSheet sheet= out.getSheet(intSheetIndex);

        WritableImage image= new WritableImage(intCol, intRow, intWidth, intHeight, file);
        sheet.addImage(image);
    }
    //----------------------------------------------------------------------------------------------------
    //Them anh toi
    //Added Date: 20-06-2008
    //----------------------------------------------------------------------------------------------------
    public void insertImage(File file, int intCol, int intRow, int intWidth, int intHeight) throws Exception {
        insertImage(0, file, intCol, intRow, intWidth, intHeight);
    }
    //----------------------------------------------------------------------------------------------------
    //Them anh toi
    //Added Date: 20-06-2008
    //----------------------------------------------------------------------------------------------------
    public void insertImage(int intSheetIndex, String strFilePath, int intCol, int intRow, int intWidth, int intHeight) throws Exception {
        File file= new File(strFilePath);
        if (!file.isFile()) throw new Exception("The image file doesn't exists");

        insertImage(intSheetIndex, file, intCol, intRow, intWidth, intHeight);
    }
    //----------------------------------------------------------------------------------------------------
    //Them anh toi
    //Added Date: 20-06-2008
    //----------------------------------------------------------------------------------------------------
    public void insertImage(String strFilePath, int intCol, int intRow, int intWidth, int intHeight) throws Exception {
        insertImage(1, strFilePath, intCol, intRow, intWidth, intHeight);
    }
    //----------------------------------------------------------------------------------------------------
    //Them anh toi
    //Added Date: 20-06-2008
    //----------------------------------------------------------------------------------------------------
    public void insertImage(int intSheetIndex, byte[] imageData, int intCol, int intRow, int intWidth, int intHeight) throws Exception {
        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
        WritableSheet sheet= out.getSheet(intSheetIndex);

        WritableImage image= new WritableImage(intCol, intRow, intWidth, intHeight, imageData);
        sheet.addImage(image);
    }
    //----------------------------------------------------------------------------------------------------
    //Them anh toi
    //Added Date: 20-06-2008
    //----------------------------------------------------------------------------------------------------
    public void insertImage(byte[] imageData, int intCol, int intRow, int intWidth, int intHeight) throws Exception {
        insertImage(0, imageData, intCol, intRow, intWidth, intHeight);
    }
    //----------------------------------------------------------------------------------------------------
    public void unInit() throws IOException, WriteException
    {
        if(out != null)
        {
            out.removeSheet(out.getNumberOfSheets() - 1);
            out.write();
            out.close();
            out = null;
        }
        for(int inti = 0; inti < vtTemplate.size(); inti++)
        {
            Template template = (Template)vtTemplate.get(inti);
            template.unInit();
        }

        if(hstParam != null)
        {
            hstParam.clear();
            hstParam = null;
        }
        vtRow.clear();
        vtRow = null;
        sheet = null;
        if (in!=null) in.close();
    }
    //--------------------------------------------------------------------------------------------
     //Thiet lap Title  cho sheet
     //Added Date: 20-06-2008
     //--------------------------------------------------------------------------------------------
     public void setSheetName( String strSheetName)  throws WriteException, Exception {
         setSheetName(1, strSheetName) ;
     }
     //--------------------------------------------------------------------------------------------
     //Thiet lap GT cho cell
     //Added Date: 20-06-2008
     //--------------------------------------------------------------------------------------------
     public void setSheetName(int intSheetIndex, String strSheetName)  throws WriteException, Exception {
         intSheetIndex--;
         if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
         if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
         WritableSheet sheet= out.getSheet(intSheetIndex);

         sheet.setName(strSheetName);
    }
    //--------------------------------------------------------------------------------------------------------------
    //Them template moi
    //--------------------------------------------------------------------------------------------------------------
    public void addTemplate(String strTitle, int intIndex, Vector vtRs)
    {
        intIndex--;

        out.copySheet(0, strTitle, intIndex);
        WritableSheet sheet= out.getSheet(intIndex);

        for(int inti=0; inti< vtTemplate.size(); inti++)
        {
            Template template= (Template)vtTemplate.get(inti);
            ResultSet rs= (ResultSet)vtRs.get(inti);

            template.setSheet(sheet);
            template.setResultSet(rs);
        }

        vtRS= vtRs;

        Template template= (Template)vtTemplate.get(0);

        /*if (intRow!=0)
        {
          for (int inti = intRow - 2; inti >= template.getDetailY(); inti--) {
            sheet.removeRow(inti);
          }
          intRow = 0;
        }*/
    }
    //--------------------------------------------------------------------------------------------------------------
    //TL template
    //--------------------------------------------------------------------------------------------------------------
    public void setTemplate(int intIndex, Vector vtRs)
    {
        intIndex--;

        WritableSheet sheet= out.getSheet(intIndex);

        for(int inti=0; inti< vtTemplate.size(); inti++)
        {
            Template template= (Template)vtTemplate.get(inti);
            ResultSet rs= (ResultSet)vtRs.get(inti);

            template.setSheet(sheet);
            template.setResultSet(rs);
        }

        vtRS= vtRs;

        Template template= (Template)vtTemplate.get(0);
    }
    //----------------------------------------------------------------------------------------------------
    //Thiet lap chieu cao hang
    //Added Date: 5-04-2009
    //----------------------------------------------------------------------------------------------------
    public void setRowHeight(int intSheetIndex, int intFromRow, int intToRow, int intHeight) throws Exception {
        intSheetIndex--;
        intFromRow--;
        intToRow--;
        if (intSheetIndex <0 ) throw new Exception("Sheet index must to be equal or greater than 0");
        if (intSheetIndex >= out.getNumberOfSheets() ) throw new Exception("Sheet index must to be less than number of sheets");
        WritableSheet sheet= out.getSheet(intSheetIndex);

        for(int inti= intFromRow; inti<=intToRow; inti++)
        {
           sheet.setRowView(inti, intHeight);
        }
    }
    //----------------------------------------------------------------------------------------------------
    //Thiet lap chieu cao hang
    //Added Date: 5-04-2009
    //----------------------------------------------------------------------------------------------------
    public void setRowHeight( int intFromRow, int intToRow, int intHeight) throws Exception {
       setRowHeight(1, intFromRow, intToRow, intHeight);
    }
}
