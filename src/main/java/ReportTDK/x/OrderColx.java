package ReportTDK.x;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */

class OrderColx
{

    int intOrderCol;
    int intStart;
    int intStep;
    int intIsReset;
    int intColToMerge;

    public OrderColx(int intOrderCol, int intStart, int intStep, int intIsReset, int intColToMerge)
    {
        this.intOrderCol = intOrderCol;
        this.intStart = intStart;
        this.intStep = intStep;
        this.intIsReset = intIsReset;
        this.intColToMerge = intColToMerge;
    }
}
