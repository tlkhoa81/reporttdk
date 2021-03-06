package ReportTDK.x;

import java.util.Vector;
import java.sql.ResultSet;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */

class Templatex
{

    private Sheet sheet;
    private int intSheetIndex;
    private int intDetailX;
    private int intDetailY;
    private int intColCount;
    private int intTemplateHeight;
    private int intHeaderPos;
    private int intHeaderHeight;
    private Vector vtGroupList;
    private Vector vtDetailList;
    private Totalx total;
    private Vector vtColToMerge;
    private OrderColx orderCol;

    public Templatex(Sheet sheet, int intDetailX, int intDetailY, int intColCount, int intTemplateHeight, int intHeaderPos, int intHeaderHeight,
            Vector vtGroupList, Vector vtDetailList, Totalx total, Vector vtColToMerge, OrderColx orderCol)
    {
        this.sheet = null;
        intSheetIndex = 0;
        this.intDetailX = 0;
        this.intDetailY = 0;
        this.intColCount = 0;
        this.intTemplateHeight = 0;
        this.intHeaderPos = 0;
        this.intHeaderHeight = 0;
        this.vtGroupList = null;
        this.vtDetailList = null;
        this.total = null;
        this.vtColToMerge = null;
        this.orderCol = null;
        this.sheet = sheet;
        this.intDetailX = intDetailX;
        this.intDetailY = intDetailY;
        this.intColCount = intColCount;
        this.intTemplateHeight = intTemplateHeight;
        this.intHeaderPos = intHeaderPos;
        this.intHeaderHeight = intHeaderHeight;
        this.vtGroupList = vtGroupList;
        this.vtDetailList = vtDetailList;
        this.total = total;
        this.vtColToMerge = vtColToMerge;
        this.orderCol = orderCol;

    }
    public Templatex(Templatex template, int intSheetIndex, Sheet sheet)
    {
        this.sheet = sheet;
        this.intSheetIndex = intSheetIndex;
        this.intDetailX = template.getDetailX();
        this.intDetailY = template.getDetailY();
        this.intColCount = template.getColCount();
        this.intTemplateHeight = template.getTemplateHeight();
        this.intHeaderPos = template.getHeaderPos();
        this.intHeaderHeight = template.getHeaderHeight();
        this.vtGroupList = template.getGroupList();
        this.vtDetailList = template.getDetailList();
        this.total = template.getTotal();
        this.vtColToMerge = template.getColToMerge();
        this.orderCol = template.getOrderCol();
        this.orderCol = orderCol;
    }

    public void setSheetIndex(int intSheetIndex)
    {
        this.intSheetIndex = intSheetIndex;
    }

    public int getSheetIndex()
    {
        return intSheetIndex;
    }

    public Sheet getSheet()
    {
        return sheet;
    }

    public int getDetailX()
    {
        return intDetailX;
    }

    public int getDetailY()
    {
        return intDetailY;
    }

    public void setDetailY(int intDetailY)
    {
        int intRowCount = intDetailY - this.intDetailY;
        if(total != null)
        {
            total.setTemplateRow1(total.getTemplateRow1() + intRowCount);
        }
        for(int inti = 0; inti < vtGroupList.size(); inti++)
        {
            Groupx group = (Groupx)vtGroupList.get(inti);
            group.intCRow += intRowCount;
        }

        this.intDetailY = intDetailY;
    }

    public int getColCount()
    {
        return intColCount;
    }

    public void setColCount(int intColCount)
    {
        this.intColCount = intColCount;
    }

    public int getTemplateHeight()
    {
        return intTemplateHeight;
    }

    public int getHeaderPos()
    {
        return intHeaderPos;
    }

    public int getHeaderHeight()
    {
        return intHeaderHeight;
    }

    public Vector getGroupList()
    {
        return vtGroupList;
    }

    public Vector getDetailList()
    {
        return vtDetailList;
    }

    public Totalx getTotal()
    {
        return total;
    }

    public Vector getColToMerge()
    {
        return vtColToMerge;
    }

    public OrderColx getOrderCol()
    {
        return orderCol;
    }
    public void setResultSet(ResultSet rs)
    {
        for (int inti=0; inti < vtGroupList.size(); inti++)
        {
            Groupx group = (Groupx)vtGroupList.get(inti);
            group.setResultSet(rs);
        }
        for (int inti=0; inti < vtDetailList.size(); inti++)
        {
            Detailx detail = (Detailx)vtDetailList.get(inti);
            detail.setResultSet(rs);
        }
        if (total!= null)  total.setResultSet(rs);
    }
    public void setSheet(Sheet sheet)
    {
        this.sheet= sheet;
        for (int inti=0; inti < vtGroupList.size(); inti++)
        {
            Groupx group = (Groupx)vtGroupList.get(inti);
            group.setSheet(sheet);
        }
        for (int inti=0; inti < vtDetailList.size(); inti++)
        {
            Detailx detail = (Detailx)vtDetailList.get(inti);
            detail.setSheet(sheet);
        }
        if (total!= null) total.setSheet(sheet);
    }
    public void unInit()
    {
        for(int intj = 0; intj < vtGroupList.size(); intj++)
        {
            Groupx group = (Groupx)vtGroupList.get(intj);
            group.unInit();
            group = null;
        }

        vtGroupList.clear();
        for(int intj = 0; intj < vtDetailList.size(); intj++)
        {
            Detailx detail = (Detailx)vtDetailList.get(intj);
            detail.unInit();
            detail = null;
        }

        vtDetailList.clear();
        if(total != null)
        {
            total.unInit();
        }
        total = null;
        vtColToMerge.clear();
        vtColToMerge = null;
    }
}
