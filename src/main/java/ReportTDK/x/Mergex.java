package ReportTDK.x;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */
class Mergex
{

    private int intC1;
    private int intR1;
    private int intC2;
    private int intR2;

    public Mergex(int intR1, int intR2, int intC1, int intC2)
    {
        this.intC1 = intC1;
        this.intR1 = intR1;
        this.intC2 = intC2;
        this.intR2 = intR2;
    }

    public int getC1()
    {
        return intC1;
    }

    public int getR1()
    {
        return intR1;
    }

    public int getC2()
    {
        return intC2;
    }

    public int getR2()
    {
        return intR2;
    }

    public void prepairMerge(int intCol)
    {
        if(intC1 >= intCol)
        {
            intC1++;
        }
        if(intC2 >= intCol)
        {
            intC2++;
        }
    }
}
