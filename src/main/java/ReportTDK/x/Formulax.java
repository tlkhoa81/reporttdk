package ReportTDK.x;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Hashtable;
import java.util.Vector;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */
class Formulax
{

    private CellStyle cf;
    public String strFormula;
    public String strFormula1;
    private int intCol;
    private int intRow;
    private Vector vtLoc;
    private Vector vtFieldList;

    public Formulax(Cell c, String strFormula)
    {
        cf = null;
        this.strFormula = "";
        strFormula1 = "";
        intCol = 0;
        intRow = 0;
        vtLoc = null;
        vtFieldList = null;
        cf = c.getCellStyle();
        intCol = c.getColumnIndex();
        intRow = c.getRowIndex();
        vtLoc = new Vector();
        vtFieldList = new Vector();
        parseFormula(strFormula);
    }

    private void parseFormula(String strFormula)
    {
        int intn = -1;
        do
        {
            int intm = strFormula.indexOf("${", intn + 1);
            if(intm < 0)
            {
                break;
            }
            intn = strFormula.indexOf("}", intm);
            if(intn < 0)
            {
                break;
            }
            String strFieldName = strFormula.substring(intm + 2, intn);
            vtFieldList.add(strFieldName);
        } while(true);
        if(vtFieldList.size() != 0)
        {
            strFormula1 = strFormula;
            strFormula = "";
        } else
        {
            for(int inti = 0; inti < strFormula.length(); inti++)
            {
                String strCol = "";
                char ch1 = strFormula.charAt(inti);
                if(!Character.isLetter(ch1))
                {
                    continue;
                }
                inti++;
                String strRow = "";
                char ch2 = strFormula.charAt(inti);
                if(!Character.isDigit(ch2))
                {
                    inti--;
                    continue;
                }
                int intk = inti - 1;
                do
                {
                    if(!Character.isLetter(strFormula.charAt(intk)))
                    {
                        break;
                    }
                    strCol = strFormula.charAt(intk) + strCol;
                } while(--intk >= 0);
                do
                {
                    if(!Character.isDigit(ch2) || inti >= strFormula.length())
                    {
                        break;
                    }
                    strRow = strRow + ch2;
                    if(++inti < strFormula.length())
                    {
                        ch2 = strFormula.charAt(inti);
                    }
                } while(true);
                if(Integer.parseInt(strRow) - 1 == intRow)
                {
                    strFormula = Groupx.replace(strFormula, "" + strCol + strRow, strCol + "#i");
                } else
                {
                    strFormula = Groupx.replace(strFormula, "" + strCol + strRow, strCol + strRow);
                }
            }

            for(int inti = strFormula.indexOf("$"); inti != -1; inti = strFormula.indexOf("$"))
            {
                int intk;
                for(intk = inti - 1; Character.isLetter(strFormula.charAt(intk)) && intk >= 0; intk--) { }
                intk++;
                String strCol = strFormula.substring(intk, inti).trim();
                int intj = inti + 1;
                String strRow = "";
                char ch2 = strFormula.charAt(intj);
                do
                {
                    if(!Character.isDigit(ch2) || intj >= strFormula.length())
                    {
                        break;
                    }
                    strRow = strRow + ch2;
                    if(++intj < strFormula.length())
                    {
                        ch2 = strFormula.charAt(intj);
                    }
                } while(true);
                strFormula = Groupx.replace(strFormula, strCol + "$" + strRow, strCol + "#j:" + strCol + "#k");
            }

            for(int inti = 0; inti + 1 < strFormula.length(); inti++)
            {
                if(strFormula.substring(inti, inti + 2).equals("#i"))
                {
                    Locx loc = new Locx("i", inti);
                    vtLoc.add(loc);
                    continue;
                }
                if(strFormula.substring(inti, inti + 2).equals("#j"))
                {
                    Locx loc = new Locx("j", inti);
                    vtLoc.add(loc);
                    continue;
                }
                if(strFormula.substring(inti, inti + 2).equals("#k"))
                {
                    Locx loc = new Locx("k", inti);
                    vtLoc.add(loc);
                }
            }

        }
        this.strFormula = strFormula;
    }

    public String getFormula(ResultSet rs)
        throws SQLException
    {
        String strFormula1 = this.strFormula1;
        int intFieldCount = vtFieldList.size();
        if(!strFormula1.equals("") && intFieldCount > 0)
        {
            for(int inti = 0; inti < intFieldCount; inti++)
            {
                String strFieldName = vtFieldList.get(inti).toString();
                String strFieldValue = "" + rs.getString(strFieldName);
                if(strFieldValue.equalsIgnoreCase("null"))
                {
                    strFieldValue = "";
                }
                strFormula1 = Groupx.replace(strFormula1, "${" + strFieldName + "}", strFieldValue);
            }

        }
        return strFormula1;
    }

    public String getFormula(int intRow)
        throws SQLException
    {
        Hashtable hstValue = new Hashtable();
        hstValue.put("i", "" + intRow);
        String strFormula1 = "";
        int intn = -2;
        for(int inti = 0; inti < vtLoc.size(); inti++)
        {
            int intm = intn + 2;
            Locx loc = (Locx)vtLoc.get(inti);
            intn = loc.intLoc;
            strFormula1 = strFormula1 + strFormula.substring(intm, intn) + hstValue.get(loc.strKey);
        }

        strFormula1 = strFormula1 + strFormula.substring(intn + 2);
        hstValue.clear();
        return strFormula1;
    }

    public String getFormula(int intRow1, int intRow2)
        throws SQLException
    {
        Hashtable hstValue = new Hashtable();
        hstValue.put("i", "" + (intRow1 - 1));
        hstValue.put("j", "" + intRow1);
        hstValue.put("k", "" + intRow2);
        String strFormula1 = "";
        int intn = -2;
        for(int inti = 0; inti < vtLoc.size(); inti++)
        {
            int intm = intn + 2;
            Locx loc = (Locx)vtLoc.get(inti);
            intn = loc.intLoc;
            strFormula1 = strFormula1 + strFormula.substring(intm, intn) + hstValue.get(loc.strKey);
        }

        strFormula1 = strFormula1 + strFormula.substring(intn + 2);
        hstValue.clear();
        return strFormula1;
    }

    public String getFormula(int intRow, int intRow1, int intRow2)
        throws SQLException
    {
        Hashtable hstValue = new Hashtable();
        hstValue.put("i", "" + intRow);
        hstValue.put("j", "" + intRow1);
        hstValue.put("k", "" + intRow2);
        String strFormula1 = "";
        int intn = -2;
        for(int inti = 0; inti < vtLoc.size(); inti++)
        {
            int intm = intn + 2;
            Locx loc = (Locx)vtLoc.get(inti);
            intn = loc.intLoc;
            strFormula1 = strFormula1 + strFormula.substring(intm, intn) + hstValue.get(loc.strKey);
        }

        strFormula1 = strFormula1 + strFormula.substring(intn + 2);
        hstValue.clear();
        return strFormula1;
    }

    public CellStyle getCellFormat()
    {
        return cf;
    }

    public int getC()
    {
        return intCol;
    }

    public void prepairFormula(Sheet sheet, int intCol)
    {
        strFormula = prepairFormula(sheet, intCol, strFormula);
        vtLoc.clear();
        for(int inti = 0; inti + 1 < strFormula.length(); inti++)
        {
            if(strFormula.substring(inti, inti + 2).equals("#i"))
            {
                Locx loc = new Locx("i", inti);
                vtLoc.add(loc);
                continue;
            }
            if(strFormula.substring(inti, inti + 2).equals("#j"))
            {
                Locx loc = new Locx("j", inti);
                vtLoc.add(loc);
                continue;
            }
            if(strFormula.substring(inti, inti + 2).equals("#k"))
            {
                Locx loc = new Locx("k", inti);
                vtLoc.add(loc);
            }
        }

        strFormula1 = prepairFormula(sheet, intCol, strFormula1);
    }

    private String prepairFormula(Sheet sheet, int intCol, String strFormula)
    {
        String strFormula1 = strFormula;
        Cell cell = null;
        int inti = 0;
        int intj = 0;
        for(intj = strFormula1.indexOf("#"); intj != -1; intj = strFormula1.indexOf("#", intj + 1))
        {
            String strColName = "";
            inti = intj - 1;
            for(char ch1 = strFormula1.charAt(inti); Character.isLetter(ch1); ch1 = strFormula1.charAt(inti))
            {
                strColName = ch1 + strColName;
                inti--;
            }

            inti++;
            if(strColName.equals(""))
            {
                continue;
            }

            int intCol1 = getColumnIndex(strColName);
            if(intCol1 >= intCol)
            {
                String strColName1 = increaseColName(strColName);
                strFormula1 = strFormula1.substring(0, inti) + strColName1 + strFormula1.substring(intj);
                intj += strColName1.length() - strColName.length();
            }
        }

        return strFormula1;
    }

    private int getColumnIndex(String strColName)
    {
       int intColIndex= (int)strColName.charAt(0) - 65;

       return intColIndex;
    }
    private String increaseColName(String strColName)
    {
        String strColName1 = strColName;
        char ch1 = strColName1.charAt(strColName1.length() - 1);
        if(ch1 == 'Z')
        {
            ch1 = 'A';
        } else
        {
            ch1 = (char)(1 + (byte)ch1);
        }
        if(ch1 == 'A')
        {
            if(strColName1.length() > 1)
            {
                strColName1 = increaseColName(strColName1.substring(0, strColName1.length() - 1)) + ch1;
            } else
            {
                strColName1 = "A" + ch1;
            }
        } else
        {
            strColName1 = strColName1.substring(0, strColName1.length() - 1) + ch1;
        }
        return strColName1;
    }
}
