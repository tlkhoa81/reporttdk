package ReportTDK;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */

class Loc
{

    public int intLoc;
    public String strKey;
    int intSheetIndex;
    public int intCol;
    public int intRow;

    public Loc(String strKey, int intLoc)
    {
        this.strKey = strKey;
        this.intLoc = intLoc;
    }

    public Loc(int intSheetIndex, int intCol, int intRow)
    {
        this.intSheetIndex = intSheetIndex;
        this.intCol = intCol;
        this.intRow = intRow;
    }
}
