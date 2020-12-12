package ReportTDK.x;

import java.io.InputStream;
import java.io.Reader;
import java.math.BigDecimal;
import java.net.URL;
import java.sql.*;
import java.util.*;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>

 * @version 2.0
 */
class VResultSetx implements ResultSet
{
    private Vector r;
    private Vector vtRow;
    private int intCurrentRow;
    private Hashtable hstColumns;

    public VResultSetx(Vector rows)
    {
        vtRow = null;
        hstColumns = null;
        r = rows;
        intCurrentRow = -1;
    }

    public void setColumnNames(String arColumns[])
    {
        hstColumns = new Hashtable();
        for(int inti = 0; inti < arColumns.length; inti++)
        {
            String strColumnName = arColumns[inti].trim();
            hstColumns.put(strColumnName, String.valueOf(inti));
        }

    }

    public boolean next()
        throws SQLException
    {
        intCurrentRow++;
        if(intCurrentRow >= r.size())
        {
            afterLast();
            return false;
        } else
        {
            vtRow = (Vector)r.get(intCurrentRow);
            return true;
        }
    }

    public void beforeFirst()
        throws SQLException
    {
        intCurrentRow = -1;
    }

    public void afterLast()
        throws SQLException
    {
        intCurrentRow = r.size();
    }

    public boolean isBeforeFirst()
    {
        return intCurrentRow < 0;
    }

    public boolean isAfterLast()
    {
        return intCurrentRow >= r.size();
    }

    public String getString(int intColumnIndex)
        throws SQLException
    {
        if(isBeforeFirst() || isAfterLast())
        {
            throw new SQLException("Invalid cursor position!");
        } else
        {
            return vtRow.get(intColumnIndex - 1).toString();
        }
    }

    public int getConcurrency()
        throws SQLException
    {
        return 0;
    }

    public int getFetchDirection()
        throws SQLException
    {
        return 0;
    }

    public int getFetchSize()
        throws SQLException
    {
        return 0;
    }

    public int getRow()
        throws SQLException
    {
        return intCurrentRow;
    }

    public int getType()
        throws SQLException
    {
        return 0;
    }

    public void cancelRowUpdates()
        throws SQLException
    {
    }

    public void clearWarnings()
        throws SQLException
    {
    }

    public void close()
        throws SQLException
    {
        if(hstColumns != null)
        {
            hstColumns.clear();
        }
    }

    public void deleteRow()
        throws SQLException
    {
    }

    public void insertRow()
        throws SQLException
    {
    }

    public void moveTointCurrentRow()
        throws SQLException
    {
    }

    public void moveToInsertRow()
        throws SQLException
    {
    }

    public void refreshRow()
        throws SQLException
    {
    }

    public void updateRow()
        throws SQLException
    {
    }

    public boolean first()
        throws SQLException
    {
        intCurrentRow = 0;
        vtRow = (Vector)r.get(intCurrentRow);
        return true;
    }

    public boolean isFirst()
        throws SQLException
    {
        return intCurrentRow == 0;
    }

    public boolean isLast()
        throws SQLException
    {
        return intCurrentRow == r.size() - 1;
    }

    public boolean last()
        throws SQLException
    {
        intCurrentRow = r.size() - 1;
        vtRow = (Vector)r.get(intCurrentRow);
        return true;
    }

    public boolean previous()
        throws SQLException
    {
        if(intCurrentRow == 0)
        {
            return false;
        } else
        {
            intCurrentRow--;
            vtRow = (Vector)r.get(intCurrentRow);
            return true;
        }
    }

    public boolean rowDeleted()
        throws SQLException
    {
        return false;
    }

    public boolean rowInserted()
        throws SQLException
    {
        return false;
    }

    public boolean rowUpdated()
        throws SQLException
    {
        return false;
    }

    public boolean wasNull()
        throws SQLException
    {
        return false;
    }

    public byte getByte(int intColumnIndex)
        throws SQLException
    {
        return Byte.parseByte(vtRow.get(intColumnIndex).toString());
    }

    public double getDouble(int intColumnIndex)
        throws SQLException
    {
        return Double.parseDouble(vtRow.get(intColumnIndex).toString());
    }

    public float getFloat(int intColumnIndex)
        throws SQLException
    {
        return Float.parseFloat(vtRow.get(intColumnIndex).toString());
    }

    public int getInt(int intColumnIndex)
        throws SQLException
    {
        return Integer.parseInt(vtRow.get(intColumnIndex).toString());
    }

    public long getLong(int intColumnIndex)
        throws SQLException
    {
        return Long.parseLong(vtRow.get(intColumnIndex).toString());
    }

    public short getShort(int intColumnIndex)
        throws SQLException
    {
        return Short.parseShort(vtRow.get(intColumnIndex).toString());
    }

    public void setFetchDirection(int i)
        throws SQLException
    {
    }

    public void setFetchSize(int i)
        throws SQLException
    {
    }

    public void updateNull(int i)
        throws SQLException
    {
    }

    public boolean absolute(int row)
        throws SQLException
    {
        return false;
    }

    public boolean getBoolean(int intColumnIndex)
        throws SQLException
    {
        return false;
    }

    public boolean relative(int rows)
        throws SQLException
    {
        return false;
    }

    public byte[] getBytes(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public void updateByte(int i, byte byte0)
        throws SQLException
    {
    }

    public void updateDouble(int i, double d)
        throws SQLException
    {
    }

    public void updateFloat(int i, float f)
        throws SQLException
    {
    }

    public void updateInt(int i, int j)
        throws SQLException
    {
    }

    public void updateLong(int i, long l)
        throws SQLException
    {
    }

    public void updateShort(int i, short word0)
        throws SQLException
    {
    }

    public void updateBoolean(int i, boolean flag)
        throws SQLException
    {
    }

    public void updateBytes(int i, byte abyte0[])
        throws SQLException
    {
    }

    public InputStream getAsciiStream(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public InputStream getBinaryStream(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public InputStream getUnicodeStream(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public void updateAsciiStream(int i, InputStream inputstream, int j)
        throws SQLException
    {
    }

    public void updateBinaryStream(int i, InputStream inputstream, int j)
        throws SQLException
    {
    }

    public Reader getCharacterStream(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public void updateCharacterStream(int i, Reader reader, int j)
        throws SQLException
    {
    }

    public Object getObject(int intColumnIndex)
        throws SQLException
    {
        return vtRow.get(intColumnIndex);
    }

    public void updateObject(int i, Object obj)
        throws SQLException
    {
    }

    public void updateObject(int i, Object obj, int j)
        throws SQLException
    {
    }

    public String getCursorName()
        throws SQLException
    {
        return "";
    }

    public void updateString(int i, String s)
        throws SQLException
    {
    }

    public byte getByte(String strColumnName)
        throws SQLException
    {
        return 0;
    }

    public double getDouble(String strColumnName)
        throws SQLException
    {
        return 0.0D;
    }

    public float getFloat(String strColumnName)
        throws SQLException
    {
        return 0.0F;
    }

    public int findColumn(String strColumnName)
        throws SQLException
    {
        return !hstColumns.containsKey(strColumnName) ? 0 : 1;
    }

    public int getInt(String strColumnName)
        throws SQLException
    {
        return 0;
    }

    public long getLong(String strColumnName)
        throws SQLException
    {
        return 0L;
    }

    public short getShort(String strColumnName)
        throws SQLException
    {
        return 0;
    }

    public void updateNull(String s)
        throws SQLException
    {
    }

    public boolean getBoolean(String strColumnName)
        throws SQLException
    {
        return false;
    }

    public byte[] getBytes(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public void updateByte(String s, byte byte0)
        throws SQLException
    {
    }

    public void updateDouble(String s, double d)
        throws SQLException
    {
    }

    public void updateFloat(String s, float f)
        throws SQLException
    {
    }

    public void updateInt(String s, int i)
        throws SQLException
    {
    }

    public void updateLong(String s, long l)
        throws SQLException
    {
    }

    public void updateShort(String s, short word0)
        throws SQLException
    {
    }

    public void updateBoolean(String s, boolean flag)
        throws SQLException
    {
    }

    public void updateBytes(String s, byte abyte0[])
        throws SQLException
    {
    }

    public BigDecimal getBigDecimal(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public BigDecimal getBigDecimal(int intColumnIndex, int scale)
        throws SQLException
    {
        return null;
    }

    public void updateBigDecimal(int i, BigDecimal bigdecimal)
        throws SQLException
    {
    }

    public URL getURL(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public Array getArray(int i)
        throws SQLException
    {
        return null;
    }

    public void updateArray(int i, Array array)
        throws SQLException
    {
    }

    public Blob getBlob(int i)
        throws SQLException
    {
        return null;
    }

    public void updateBlob(int i, Blob blob)
        throws SQLException
    {
    }

    public Clob getClob(int i)
        throws SQLException
    {
        return null;
    }

    public void updateClob(int i, Clob clob)
        throws SQLException
    {
    }

    public Ref getRef(int i)
        throws SQLException
    {
        return null;
    }

    public void updateRef(int i, Ref ref)
        throws SQLException
    {
    }

    public ResultSetMetaData getMetaData()
        throws SQLException
    {
        return null;
    }

    public SQLWarning getWarnings()
        throws SQLException
    {
        return null;
    }

    public Statement getStatement()
        throws SQLException
    {
        return null;
    }

    public Time getTime(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public void updateTime(int i, Time time)
        throws SQLException
    {
    }

    public Timestamp getTimestamp(int intColumnIndex)
        throws SQLException
    {
        return null;
    }

    public void updateTimestamp(int i, Timestamp timestamp)
        throws SQLException
    {
    }

    public InputStream getAsciiStream(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public InputStream getBinaryStream(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public InputStream getUnicodeStream(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public void updateAsciiStream(String s, InputStream inputstream, int i)
        throws SQLException
    {
    }

    public void updateBinaryStream(String s, InputStream inputstream, int i)
        throws SQLException
    {
    }

    public Reader getCharacterStream(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public void updateCharacterStream(String s, Reader reader1, int i)
        throws SQLException
    {
    }

    public Object getObject(String strColumnName) throws SQLException
    {
        Object objValue= null;
        int intColumnIndex= -1;

        if (hstColumns.containsKey(strColumnName))
        {
           intColumnIndex = Integer.parseInt(hstColumns.get(strColumnName).toString());
           objValue= vtRow.get(intColumnIndex);
        }
        return objValue;
    }

    public void updateObject(String s, Object obj)
        throws SQLException
    {
    }

    public void updateObject(String s, Object obj, int i)
        throws SQLException
    {
    }

    public Object getObject(int i, Map map)
        throws SQLException
    {
        return null;
    }

    public String getString(String strColumnName)
        throws SQLException
    {
        if(isBeforeFirst() || isAfterLast())
        {
            throw new SQLException("Invalid cursor position!");
        } else
        {
            String strColumnIndex = (String)hstColumns.get(strColumnName);
            int intColumnIndex = Integer.parseInt(strColumnIndex);
            Vector row = (Vector)r.elementAt(intCurrentRow);
            return row.get(intColumnIndex).toString();
        }
    }

    public void updateString(String s, String s1)
        throws SQLException
    {
    }

    public BigDecimal getBigDecimal(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public BigDecimal getBigDecimal(String strColumnName, int scale)
        throws SQLException
    {
        return null;
    }

    public void updateBigDecimal(String s, BigDecimal bigdecimal)
        throws SQLException
    {
    }

    public URL getURL(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public Array getArray(String colName)
        throws SQLException
    {
        return null;
    }

    public void updateArray(String s, Array array)
        throws SQLException
    {
    }

    public Blob getBlob(String colName)
        throws SQLException
    {
        return null;
    }

    public void updateBlob(String s, Blob blob)
        throws SQLException
    {
    }

    public Clob getClob(String colName)
        throws SQLException
    {
        return null;
    }

    public void updateClob(String s, Clob clob)
        throws SQLException
    {
    }

    public Ref getRef(String colName)
        throws SQLException
    {
        return null;
    }

    public void updateRef(String s, Ref ref)
        throws SQLException
    {
    }

    public Time getTime(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public void updateTime(String s, Time time)
        throws SQLException
    {
    }

    public Time getTime(int intColumnIndex, Calendar cal)
        throws SQLException
    {
        return null;
    }

    public Timestamp getTimestamp(String strColumnName)
        throws SQLException
    {
        return null;
    }

    public void updateTimestamp(String s, Timestamp timestamp)
        throws SQLException
    {
    }

    public Timestamp getTimestamp(int intColumnIndex, Calendar cal)
        throws SQLException
    {
        return null;
    }

    public Object getObject(String colName, Map map)
        throws SQLException
    {
        return null;
    }

    public Time getTime(String strColumnName, Calendar cal)
        throws SQLException
    {
        return null;
    }

    public Timestamp getTimestamp(String strColumnName, Calendar cal)
        throws SQLException
    {
        return null;
    }

    public void moveToCurrentRow()
        throws SQLException
    {
    }

  public java.sql.Date getDate(int columnIndex) throws SQLException {
    return null;
  }

  public void updateDate(int columnIndex, java.sql.Date x) throws SQLException {
  }

  public java.sql.Date getDate(String columnName) throws SQLException {
    return null;
  }

  public void updateDate(String columnName, java.sql.Date x) throws
      SQLException {
  }

  public java.sql.Date getDate(int columnIndex, Calendar cal) throws
      SQLException {
    return null;
  }

  public java.sql.Date getDate(String columnName, Calendar cal) throws
      SQLException {
    return null;
  }

public <T> T unwrap(Class<T> iface) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public boolean isWrapperFor(Class<?> iface) throws SQLException {
	// TODO Auto-generated method stub
	return false;
}

public RowId getRowId(int columnIndex) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public RowId getRowId(String columnLabel) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public void updateRowId(int columnIndex, RowId x) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateRowId(String columnLabel, RowId x) throws SQLException {
	// TODO Auto-generated method stub
	
}

public int getHoldability() throws SQLException {
	// TODO Auto-generated method stub
	return 0;
}

public boolean isClosed() throws SQLException {
	// TODO Auto-generated method stub
	return false;
}

public void updateNString(int columnIndex, String nString) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNString(String columnLabel, String nString) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNClob(int columnIndex, NClob nClob) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNClob(String columnLabel, NClob nClob) throws SQLException {
	// TODO Auto-generated method stub
	
}

public NClob getNClob(int columnIndex) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public NClob getNClob(String columnLabel) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public SQLXML getSQLXML(int columnIndex) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public SQLXML getSQLXML(String columnLabel) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public void updateSQLXML(int columnIndex, SQLXML xmlObject) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateSQLXML(String columnLabel, SQLXML xmlObject) throws SQLException {
	// TODO Auto-generated method stub
	
}

public String getNString(int columnIndex) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public String getNString(String columnLabel) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public Reader getNCharacterStream(int columnIndex) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public Reader getNCharacterStream(String columnLabel) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public void updateNCharacterStream(int columnIndex, Reader x, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNCharacterStream(String columnLabel, Reader reader, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateAsciiStream(int columnIndex, InputStream x, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateBinaryStream(int columnIndex, InputStream x, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateCharacterStream(int columnIndex, Reader x, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateAsciiStream(String columnLabel, InputStream x, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateBinaryStream(String columnLabel, InputStream x, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateCharacterStream(String columnLabel, Reader reader, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateBlob(int columnIndex, InputStream inputStream, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateBlob(String columnLabel, InputStream inputStream, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateClob(int columnIndex, Reader reader, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateClob(String columnLabel, Reader reader, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNClob(int columnIndex, Reader reader, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNClob(String columnLabel, Reader reader, long length) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNCharacterStream(int columnIndex, Reader x) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNCharacterStream(String columnLabel, Reader reader) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateAsciiStream(int columnIndex, InputStream x) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateBinaryStream(int columnIndex, InputStream x) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateCharacterStream(int columnIndex, Reader x) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateAsciiStream(String columnLabel, InputStream x) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateBinaryStream(String columnLabel, InputStream x) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateCharacterStream(String columnLabel, Reader reader) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateBlob(int columnIndex, InputStream inputStream) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateBlob(String columnLabel, InputStream inputStream) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateClob(int columnIndex, Reader reader) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateClob(String columnLabel, Reader reader) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNClob(int columnIndex, Reader reader) throws SQLException {
	// TODO Auto-generated method stub
	
}

public void updateNClob(String columnLabel, Reader reader) throws SQLException {
	// TODO Auto-generated method stub
	
}

public <T> T getObject(int columnIndex, Class<T> type) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}

public <T> T getObject(String columnLabel, Class<T> type) throws SQLException {
	// TODO Auto-generated method stub
	return null;
}
}
