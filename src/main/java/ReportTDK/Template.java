package ReportTDK;

import java.util.Vector;
import jxl.write.WritableSheet;
import java.sql.ResultSet;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */

class Template
{

    private WritableSheet sheet;
    private int intSheetIndex;
    private int intDetailX;
    private int intDetailY;
    private int intColCount;
    private int intTemplateHeight;
    private int intHeaderPos;
    private int intHeaderHeight;
    private Vector vtGroupList;
    private Vector vtDetailList;
    private Total total;
    private Vector vtColToMerge;
    private OrderCol orderCol;

    public Template(WritableSheet sheet, int intDetailX, int intDetailY, int intColCount, int intTemplateHeight, int intHeaderPos, int intHeaderHeight,
            Vector vtGroupList, Vector vtDetailList, Total total, Vector vtColToMerge, OrderCol orderCol)
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
    public Template(Template template, int intSheetIndex, WritableSheet sheet)
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

    public WritableSheet getSheet()
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
            Group group = (Group)vtGroupList.get(inti);
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

    public Total getTotal()
    {
        return total;
    }

    public Vector getColToMerge()
    {
        return vtColToMerge;
    }

    public OrderCol getOrderCol()
    {
        return orderCol;
    }
    public void setResultSet(ResultSet rs)
    {
        for (int inti=0; inti < vtGroupList.size(); inti++)
        {
            Group group = (Group)vtGroupList.get(inti);
            group.setResultSet(rs);
        }
        for (int inti=0; inti < vtDetailList.size(); inti++)
        {
            Detail detail = (Detail)vtDetailList.get(inti);
            detail.setResultSet(rs);
        }
        if (total!= null)  total.setResultSet(rs);
    }
    public void setSheet(WritableSheet sheet)
    {
        this.sheet= sheet;
        for (int inti=0; inti < vtGroupList.size(); inti++)
        {
            Group group = (Group)vtGroupList.get(inti);
            group.setSheet(sheet);
        }
        for (int inti=0; inti < vtDetailList.size(); inti++)
        {
            Detail detail = (Detail)vtDetailList.get(inti);
            detail.setSheet(sheet);
        }
        if (total!= null) total.setSheet(sheet);
    }
    public void unInit()
    {
        for(int intj = 0; intj < vtGroupList.size(); intj++)
        {
            Group group = (Group)vtGroupList.get(intj);
            group.unInit();
            group = null;
        }

        vtGroupList.clear();
        for(int intj = 0; intj < vtDetailList.size(); intj++)
        {
            Detail detail = (Detail)vtDetailList.get(intj);
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
