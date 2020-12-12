package ReportTDK.x;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */

class Locx
{

    public int intLoc;
    public String strKey;
    int intSheetIndex;
    public int intCol;
    public int intRow;

    public Locx(String strKey, int intLoc)
    {
        this.strKey = strKey;
        this.intLoc = intLoc;
    }

    public Locx(int intSheetIndex, int intCol, int intRow)
    {
        this.intSheetIndex = intSheetIndex;
        this.intCol = intCol;
        this.intRow = intRow;
    }
}
