package ReportTDK.x;

import java.util.Hashtable;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */
class Util {
   public Util() {
   }
   //-------------------------------------------------------------------------------------------------------------
   //Remove formula
   //Added Date: 04-10-2008
   //-------------------------------------------------------------------------------------------------------------
   public static void removeFormula(Sheet sheet) throws Exception {
      int intRowCount= sheet.getLastRowNum()+ 1;
      int intColCount =0;
      int intRow=0;
      while(intRow<=10)
      {
        Row row= sheet.getRow(intRow);
        if (row==null) { intRow++; continue;}
        if (intColCount< row.getLastCellNum()+ 1) intColCount = row.getLastCellNum() + 1;
        intRow++;
      }
      for (int intj=0; intj < intRowCount; intj++)
      {
         for(int inti= 0; inti < intColCount; inti++)
         {
            Row row= sheet.getRow(intj); if (row==null) continue;
            Cell cell= row.getCell(inti); if (cell==null) continue;

            if (getCellType(cell)==cell.CELL_TYPE_FORMULA) //.getClass().getName().equals("jxl.write.Formula"
            {
               String strValue= getCellValue(sheet, inti, intj);

               addCellToSheet(sheet, cell, 2, strValue, inti, intj);
            }
         }
      }
   }
   //------------------------------------------------------------------------------------------------------------------------
   //Write cell to sheet
   //------------------------------------------------------------------------------------------------------------------------
   public  static void addCellToSheet(Sheet sheet, Cell templateCell, int intCellType, String strValue, int intCol, int intRow)
   {
       CellStyle cf= templateCell.getCellStyle();

       Row row= sheet.getRow(intRow);
       if (row==null) row= sheet.createRow(intRow);
       Cell cell= row.getCell(intCol);
       if (cell==null) cell= row.createCell(intCol);

       if(strValue.equals(""))
       {
           cell.setCellType(cell.CELL_TYPE_BLANK);

           if (cf!=null) cell.setCellStyle(cf);

           return;
       }

       if (Util.isNumeric(strValue)) intCellType= Cell.CELL_TYPE_NUMERIC;
       switch(intCellType)
       {
       case Cell.CELL_TYPE_STRING: // '\0'
         cell.setCellType(cell.CELL_TYPE_STRING);
         cell.setCellValue(strValue);

         if (cf!=null) cell.setCellStyle(cf);

         break;
       case Cell.CELL_TYPE_BOOLEAN: // '\001'
         cell.setCellType(cell.CELL_TYPE_BOOLEAN);
         cell.setCellValue(strValue);

         if (cf!=null) cell.setCellStyle(cf);

         break;
       case Cell.CELL_TYPE_NUMERIC: // '\002'
           cell.setCellType(cell.CELL_TYPE_NUMERIC);
           cell.setCellValue(Double.parseDouble(strValue));

           if (cf!=null) cell.setCellStyle(cf);

           break;
       }
    }
   //------------------------------------------------------------------------------------------------------------------------
   //get Cell Value having formula
   //------------------------------------------------------------------------------------------------------------------------
   public static String getCellValue(Sheet sheet, int intCol, int intRow) throws Exception {
       String strContent= "";
       String strValue = "";

       try
       {
         if(sheet == null)
         {
             return "";
         }
         Cell cell = sheet.getRow(intRow).getCell(intCol);
         if(cell == null)
         {
             return "";
         }
         strContent = getCellValue(sheet, intCol, intRow, "");
         double dblValue;
         if(strContent.indexOf("*") != -1)
         {
            dblValue = 1.0D;
            String strCellValue = "";
            for(int inti = strContent.indexOf("*"); inti != -1; inti = strContent.indexOf("*"))
            {
                String strCell1 = strContent.substring(0, inti);
                if(isNumeric(strCell1))
                {
                    strCellValue = strCell1;
                } else
                {
                    strCellValue = Groupx.replace(getCellValue(sheet, strCell1), ",", "");
                }
                dblValue *= isNumeric(strCellValue) ? Double.parseDouble(strCellValue) : 0.0D;
                strContent = strContent.substring(inti + 1);
            }

            if(!strContent.trim().equals(""))
            {
                if(isNumeric(strContent))
                {
                    strCellValue = strContent;
                } else
                {
                    strCellValue = Groupx.replace(getCellValue(sheet, strContent), ",", "");
                }
                dblValue *= isNumeric(strCellValue) ? Double.parseDouble(strCellValue) : 0.0D;
            }
            return String.valueOf(dblValue);
         }

         if(strContent.indexOf("/") != -1)
         {
           dblValue = 1.0D;
           String strCellValue = "";
           for(int inti = strContent.indexOf("/"); inti != -1; inti = strContent.indexOf("/"))
           {
               String strCell1 = strContent.substring(0, inti);
               if(isNumeric(strCell1))
               {
                   strCellValue = strCell1;
               } else
               {
                   strCellValue = Groupx.replace(getCellValue(sheet, strCell1), ",", "");
               }
               dblValue /= isNumeric(strCellValue) ? Double.parseDouble(strCellValue) : 0.0D;
               strContent = strContent.substring(inti + 1);
           }

           if(!strContent.trim().equals(""))
           {
               if(isNumeric(strContent))
               {
                   strCellValue = strContent;
               } else
               {
                   strCellValue = Groupx.replace(getCellValue(sheet, strContent), ",", "");
               }
               dblValue /= isNumeric(strCellValue) ? Double.parseDouble(strCellValue) : 0.0D;
           }
           return String.valueOf(dblValue);
         }

         if(strContent.indexOf("+") != -1)
         {
           dblValue = 0.0D;
           String strCellValue = "";
           for(int inti = strContent.indexOf("+"); inti != -1; inti = strContent.indexOf("+"))
           {
               String strCell1 = strContent.substring(0, inti);
               if(isNumeric(strCell1))
               {
                   strCellValue = strCell1;
               } else
               {
                   strCellValue = Groupx.replace(getCellValue(sheet, strCell1), ",", "");
               }
               dblValue += isNumeric(strCellValue) ? Double.parseDouble(strCellValue) : 0.0D;
               strContent = strContent.substring(inti + 1);
           }

           if(!strContent.trim().equals(""))
           {
               if(isNumeric(strContent))
               {
                   strCellValue = strContent;
               } else
               {
                   strCellValue = Groupx.replace(getCellValue(sheet, strContent), ",", "");
               }
               dblValue += isNumeric(strCellValue) ? Double.parseDouble(strCellValue) : 0.0D;
           }
           return String.valueOf(dblValue);
         }

         int inti;
         if(strContent.indexOf("SUM") != -1 && strContent.indexOf("SUMIF") == -1)
         {
           inti = strContent.indexOf("(");
           if(inti == -1)
           {
               return strContent;
           }
           int intj = strContent.indexOf(":", inti + 1);
           if(intj == -1)
           {
               return strContent;
           }
           int intk = strContent.indexOf(")", intj + 1);
           if(intk == -1)
           {
               return strContent;
           }
           Cell cell1;
           String strCell1 = strContent.substring(inti + 1, intj).trim();
           CellReference cr= new CellReference(strCell1);
           cell1 = sheet.getRow(cr.getRow()).getCell(cr.getCol());
           if(cell1 == null)
           {
               return strContent;
           }
           int intRow1;
           int intCol1;
           Cell cell2;
           intRow1 = cell1.getRowIndex();
           intCol1 = cell1.getColumnIndex();
           String strCell2 = strContent.substring(intj + 1, intk).trim();
           CellReference cr2= new CellReference(strCell2);
           cell2 = sheet.getRow(cr2.getRow()).getCell(cr2.getCol());
           if(cell2 == null)
           {
               return strContent;
           }
           int intRow2;
           int intCol2;
           intRow2 = cell2.getRowIndex();
           intCol2 = cell2.getColumnIndex();
           if(intRow1 == intRow2 && intCol1 == intCol2 || intRow1 != intRow2 && intCol1 != intCol2)
           {
               return getCellValue(sheet, intCol1, intRow1);
           }

           if(intRow1 == intRow2)
           {
             dblValue = 0.0D;
             for(inti = intCol1; inti <= intCol2; inti += intCol2 < intCol1 ? -1 : 1)
             {
                 String strValue1 = getCellValue(sheet, inti, intRow1);
                 dblValue += isNumeric(strValue1) ? Double.parseDouble(strValue1) : 0.0D;
             }

             return String.valueOf(dblValue);
           }else
           {
             dblValue = 0.0D;
             for(inti = intRow1; inti <= intRow2; inti += intRow2 < intRow1 ? -1 : 1)
             {
                 String strValue1 = getCellValue(sheet, intCol1, inti);
                 dblValue += isNumeric(strValue1) ? Double.parseDouble(strValue1) : 0.0D;
             }

             return String.valueOf(dblValue);
           }
         }

         if(strContent.indexOf("SUMIF") != -1)
         {
           inti = strContent.indexOf("(");
           if(inti == -1)
           {
               return strContent;
           }
           int intj = strContent.indexOf(",", inti + 1);
           if(intj == -1)
           {
               return strContent;
           }
           int intk = strContent.indexOf(",", intj + 1);
           if(intk == -1)
           {
               return strContent;
           }
           int intl = strContent.indexOf(")", intk + 1);
           if(intl == -1)
           {
               return strContent;
           }
           String strCompare;
           Hashtable hstLoc1;
           Hashtable hstLoc2;
           String strRange1 = strContent.substring(inti + 1, intj);
           strCompare = strContent.substring(intj + 1, intk).replaceAll("\"", "");
           String strRange2 = strContent.substring(intk + 1, intl);
           hstLoc1 = getRangeLoc(sheet, "(" + strRange2 + ")");
           hstLoc2 = getRangeLoc(sheet, "(" + strRange1 + ")");
           if(hstLoc1 == null || hstLoc2 == null)
           {
               return strContent;
           }
           dblValue = 0.0D;
           int intRow11 = Integer.parseInt(hstLoc1.get("row1").toString());
           int intCol11 = Integer.parseInt(hstLoc1.get("col1").toString());
           int intRow12 = Integer.parseInt(hstLoc1.get("row2").toString());
           int intCol12 = Integer.parseInt(hstLoc1.get("col2").toString());
           int intRow21 = Integer.parseInt(hstLoc2.get("row1").toString());
           int intCol21 = Integer.parseInt(hstLoc2.get("col1").toString());
           int intRow22 = Integer.parseInt(hstLoc2.get("row2").toString());
           int intCol22 = Integer.parseInt(hstLoc2.get("col2").toString());
           for(inti = intRow11; inti <= intRow12; inti += intRow12 < intRow11 ? -1 : 1)
           {
               String strCompare1 = getCellValue(sheet, intCol11, inti);
               if(strCompare1.equals(strCompare))
               {
                   String strValue1 = Groupx.replace(getCellValue(sheet, intCol21, inti), ",", "");
                   dblValue += isNumeric(strValue1) ? Double.parseDouble(strValue1) : 0.0D;
               }
           }

           return String.valueOf(dblValue);
       }
       }catch(Exception ex)
       {
          return "0";
       }
       return strContent;
   }
   static public String getCellValue(Sheet sheet, int intCol, int intRow, String strDefault)
   {
       String strResult = "";

       Row row= sheet.getRow(intRow);

       if (row==null) return strDefault;
       Cell cell = row.getCell(intCol);

       if(cell == null)
       {
           return strDefault;
       }

       switch(cell.getCellType())
       {
         case Cell.CELL_TYPE_BLANK: strResult= "";
           break;
         case Cell.CELL_TYPE_BOOLEAN: boolean blnResult= cell.getBooleanCellValue();
           strResult= String.valueOf(blnResult);
           break;
         case Cell.CELL_TYPE_ERROR: strResult= String.valueOf(cell.getErrorCellValue());
           break;
         case Cell.CELL_TYPE_FORMULA: strResult= String.valueOf(cell.getStringCellValue()).trim();
           break;
         case Cell.CELL_TYPE_NUMERIC: strResult= String.valueOf(cell.getNumericCellValue());
           break;
         case Cell.CELL_TYPE_STRING: strResult= String.valueOf(cell.getStringCellValue());
           break;
       }

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
   private static Hashtable getRangeLoc(Sheet sheet, String strRange)
   {
       Hashtable hstLoc = new Hashtable();
       /*tamboint inti = strRange.indexOf("(");
       if(inti == -1)
       {
           return null;
       }
       int intj = strRange.indexOf(":", inti + 1);
       if(intj == -1)
       {
           return null;
       }
       int intk = strRange.indexOf(")", intj + 1);
       if(intk == -1)
       {
           return null;
       }
       String strCell1 = strRange.substring(inti + 1, intj).trim();
       Cell cell1 = sheet.getCell(strCell1);
       if(cell1 == null)
       {
           return null;
       }
       int intRow1 = cell1.getRow();
       int intCol1 = cell1.getColumn();
       String strCell2 = strRange.substring(intj + 1, intk).trim();
       Cell cell2 = sheet.getCell(strCell2);
       if(cell2 == null)
       {
           return null;
       } else
       {
           int intRow2 = cell2.getRow();
           int intCol2 = cell2.getColumn();
           hstLoc.put("row1", String.valueOf(intRow1));
           hstLoc.put("col1", String.valueOf(intCol1));
           hstLoc.put("row2", String.valueOf(intRow2));
           hstLoc.put("col2", String.valueOf(intCol2));
           return hstLoc;
       }*/
       return hstLoc;
    }
   public static String getCellValue(Sheet sheet, String strLoc) throws Exception {
       String strValue = "";

       /*tamboCell cell = sheet.getCell(strLoc);
       if(cell == null)
       {
           return "";
       } else
       {
           int intRow = cell.getRow();
           int intCol = cell.getColumn();
           strValue = getCellValue(sheet, intCol, intRow);

           return strValue;
       }*/
       return strValue;
    }
    //------------------------------------------------------------------------------------------------------------------------
    //Check is numeric
    //------------------------------------------------------------------------------------------------------------------------
    public static boolean isNumeric(String strInput)
    {
        double dblValue;
        try
        {
            dblValue = Double.parseDouble(strInput);
        }
        catch(Exception ex)
        {
            return false;
        }
        return true;
    }
    //------------------------------------------------------------------------------------------------------------------------
    //get Cell Type
    //------------------------------------------------------------------------------------------------------------------------
    public static int getCellType(Cell cell1)
    {
        int intCellType = 0;

        intCellType= cell1.getCellType();

        return intCellType;
    }
}
