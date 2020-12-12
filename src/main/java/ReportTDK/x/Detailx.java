package ReportTDK.x;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Vector;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */
class Detailx
{

    private Sheet sheet;
    private ResultSet rs;
    private int intTemplateRow;
    private int intColCount;
    private Cell vtTemplateCell[];
    public int vtCellType[];
    private Vector vtTemplateMerge;
    private Vector vtTemplateFormula;
    private Vector vtFieldList[];
    private String vtFieldValueList[];
    private int intDetailX;
    private int intTemplateHeight;

    public Detailx(Sheet sheet, ResultSet rs, int intTemplateRow, int intColCount, int intDetailX)
        throws SQLException, Exception
    {
        this.sheet = null;
        this.rs = null;
        this.intTemplateRow = 0;
        this.intColCount = 1;
        vtTemplateCell = null;
        vtCellType = null;
        vtTemplateMerge = null;
        vtTemplateFormula = null;
        vtFieldList = null;
        vtFieldValueList = null;
        this.intDetailX = 0;
        intTemplateHeight = 0;
        intTemplateRow--;
        this.sheet = sheet;
        this.rs = rs;
        intColCount++;
        this.intTemplateRow = intTemplateRow;
        this.intDetailX = intDetailX;
        this.intColCount = intColCount;
        intTemplateHeight = intTemplateHeight;
        vtFieldList = new Vector[intColCount];
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti] = new Vector();
        }

        vtTemplateCell = new Cell[intColCount + 1];
        vtCellType = new int[intColCount + 1];
        vtTemplateMerge = new Vector();
        vtTemplateFormula = new Vector();
        vtFieldValueList = new String[intColCount];
        parseDetailTemplate();
    }

    public int getTemplateRow()
    {
        return intTemplateRow;
    }

    public void increaseColCount()
    {
        intColCount++;
        vtFieldList = new Vector[intColCount];
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti] = new Vector();
        }

        vtTemplateCell = new Cell[intColCount + 1];
        vtCellType = new int[intColCount + 1];
        vtFieldValueList = new String[intColCount];
    }

    public void setSheet(Sheet sheet)
    {
        this.sheet = sheet;
    }

    private void parseDetailTemplate()
        throws  SQLException, Exception
    {
        for(int inti = 0; inti < intColCount; inti++)
        {
            Cell cell1 = sheet.getRow(intTemplateRow).getCell(inti + intDetailX);
            vtTemplateCell[inti] = cell1;
            vtCellType[inti] = Util.getCellType(cell1);
            if(vtCellType[inti]==Cell.CELL_TYPE_FORMULA) //cell1.toString().toLowerCase().indexOf("formula"
            {
                Formulax formula = new Formulax(cell1, cell1.getCellFormula());
                vtTemplateFormula.add(formula);
                vtFieldValueList[inti] = "";
                vtCellType[inti] = 3;
                continue;
            }
            String cv = Util.getCellValue(sheet, inti + intDetailX, intTemplateRow, "");//cell1.getStringCellValue();
            vtFieldValueList[inti] = cv;
            int intn = -1;
            do
            {
                int intm = cv.indexOf("${", intn + 1);
                if(intm < 0)
                {
                    break;
                }
                intn = cv.indexOf("}", intm);
                if(intn < 0)
                {
                    break;
                }
                String strFieldName = cv.substring(intm + 2, intn);
                if(!Groupx.FieldExists(rs, strFieldName))
                {
                    throw new Exception("Error: Field " + strFieldName + " that be needed to fill at cell(" + (inti + intDetailX + 1) + ", " + (intTemplateRow + 1) + " ) does not exists");
                }
                vtFieldList[inti].add(strFieldName);
            } while(true);
            if(vtFieldList[inti].size() > 0 && vtCellType[inti] == 0)
            {
                throw new Exception("Cell( " + (inti + 1) + ", " + (intTemplateRow + 1) + ") must to be formated as text or numeric.");
            }
        }


        //Range rl1[] = sheet.getMergedCells();

        for(int inti1 = 0; inti1 < sheet.getNumMergedRegions(); inti1++)
        {
            CellRangeAddress cra= sheet.getMergedRegion(inti1);

            //Cell cell1 = cra. ;//r1.getTopLeft();
            //Cell cell2 = r1.getBottomRight();

            if(cra.getFirstRow() == cra.getLastRow() && cra.getFirstRow() == intTemplateRow)
            {
                vtTemplateMerge.add(new Mergex(intTemplateRow, intTemplateRow, cra.getFirstColumn(), cra.getLastColumn()));
                //moibosheet.ungroupColumn(cra.getFirstColumn(), cra.getLastColumn());
                //sheet.unmergeCells(r1);
            }
        }

    }

    public void reParseDetailTemplate()
        throws  SQLException, Exception
    {
        for(int inti = 0; inti < intColCount; inti++)
        {
            Cell cell1 = sheet.getRow(intTemplateRow).getCell(inti + intDetailX);
            vtTemplateCell[inti] = cell1;
            vtCellType[inti] = Util.getCellType(cell1);
            if(cell1.toString().toLowerCase().indexOf("formula") == -1)
            {
                String cv = Util.getCellValue(sheet, inti + intDetailX,intTemplateRow, "" );//cell1.getStringCellValue();
                vtFieldValueList[inti] = cv;
                vtFieldList[inti].clear();
                int intn = -1;
                do
                {
                    int intm = cv.indexOf("${", intn + 1);
                    if(intm < 0)
                    {
                        break;
                    }
                    intn = cv.indexOf("}", intm);
                    if(intn < 0)
                    {
                        break;
                    }
                    String strFieldName = cv.substring(intm + 2, intn);
                    if(!Groupx.FieldExists(rs, strFieldName))
                    {
                        throw new Exception("Error: Field " + strFieldName + " that be needed to fill at cell(" + (inti + intDetailX + 1) + ", " + (intTemplateRow + 1) + " ) does not exists");
                    }
                    vtFieldList[inti].add(strFieldName);
                } while(true);
                if(vtFieldList[inti].size() > 0 && vtCellType[inti] == 0)
                {
                    throw new Exception("Cell( " + (inti + 1) + ", " + (intTemplateRow + 1) + ") must to be formated as text or numeric.");
                }
            } else
            {
                vtFieldValueList[inti] = "";
                vtCellType[inti] = 3;
            }
        }

    }

    public void fillDetail(int intRow)  throws SQLException
    {
        sheet.shiftRows(intRow, intRow+20, 1);
        Groupx.addMergeToSheet(sheet, vtTemplateMerge, intRow);
        for(int intx = 0; intx < intColCount; intx++)
        {
            String strValue = Groupx.getCellValueToFill(rs, vtFieldList[intx], vtFieldValueList[intx]);
            //String strValue1= sheet.getRow(3).getCell(1).getStringCellValue();

            Util.addCellToSheet(sheet, vtTemplateCell[intx], vtCellType[intx], strValue, intx + intDetailX, intRow);
        }

        fillFormula(intRow);
    }

    public void fillFormula(int intRow)  throws  SQLException
    {
        for(int inti = 0; inti < vtTemplateFormula.size(); inti++)
        {
            Formulax f = (Formulax)vtTemplateFormula.get(inti);

            Row row= sheet.getRow(intRow); if (row==null) row= sheet.createRow(intRow);
            Cell cell= row.getCell(f.getC()); if (cell==null) cell= row.createCell(f.getC());
            //Cell cell= sheet.createRow(intRow).createCell(f.getC());

            cell.setCellStyle(f.getCellFormat());
            cell.setCellFormula(f.getFormula(intRow + 1));
            //sheet.addCell(new jxl.write.Formula(f.getC(), intRow, f.getFormula(intRow + 1), f.getCellFormat()));
        }

    }

    public void prepairFormulas(int intCol)
    {
        prepairFormulas(intCol, vtTemplateFormula);
    }

    private void prepairFormulas(int intCol, Vector vtFormula)
    {
        for(int inti = 0; inti < vtFormula.size(); inti++)
        {
            Formulax formula = (Formulax)vtFormula.get(inti);
            formula.prepairFormula(sheet, intCol);
        }

    }

    public void prepairMerges(int intCol)
    {
        for(int inti = 0; inti < vtTemplateMerge.size(); inti++)
        {
            Mergex merge = (Mergex)vtTemplateMerge.get(inti);
            merge.prepairMerge(intCol);
        }

    }

    public void addFormula(int intCol, String strFormula)
    {
        Cell cell = sheet.getRow(intTemplateRow).getCell(intCol);
        Formulax formula = new Formulax(cell, strFormula);
        vtTemplateFormula.add(formula);
        vtFieldValueList[intCol] = "";
        vtCellType[intCol] = 3;
    }
    public void setResultSet(ResultSet rs)
    {
        this.rs= rs;
    }
    public void setResultSet(Sheet sheet)
    {
        this.sheet= sheet;
    }
    public void unInit()
    {
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti].clear();
        }

        vtFieldList = null;
        vtTemplateCell = null;
        vtTemplateMerge.clear();
        vtTemplateMerge = null;
        vtTemplateFormula.clear();
        vtTemplateFormula = null;
        vtFieldValueList = null;
    }
}
