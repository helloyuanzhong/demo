package com.example.demo.utils;

import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelMergedRegion {
  private Integer firstRow;
  private Integer lastRow;
  private Integer firstColumn;
  private Integer lastColumn;

  public ExcelMergedRegion(CellRangeAddress cra) {
    this.firstRow = cra.getFirstRow();
    this.lastRow = cra.getLastRow();
    this.firstColumn = cra.getFirstColumn();
    this.lastColumn = cra.getLastColumn();
  }

  public ExcelMergedRegion(
      Integer firstRow, Integer lastRow, Integer firstColumn, Integer lastColumn) {
    this.firstRow = firstRow;
    this.lastRow = lastRow;
    this.firstColumn = firstColumn;
    this.lastColumn = lastColumn;
  }

  public Integer getFirstRow() {
    return this.firstRow;
  }

  public void setFirstRow(Integer firstRow) {
    this.firstRow = firstRow;
  }

  public Integer getLastRow() {
    return this.lastRow;
  }

  public void setLastRow(Integer lastRow) {
    this.lastRow = lastRow;
  }

  public Integer getFirstColumn() {
    return this.firstColumn;
  }

  public void setFirstColumn(Integer firstColumn) {
    this.firstColumn = firstColumn;
  }

  public Integer getLastColumn() {
    return this.lastColumn;
  }

  public void setLastColumn(Integer lastColumn) {
    this.lastColumn = lastColumn;
  }

  @Override
  public int hashCode() {
    boolean prime = true;
    int result = 1;
    result = 31 * result + (this.firstColumn == null ? 0 : this.firstColumn.hashCode());
    result = 31 * result + (this.firstRow == null ? 0 : this.firstRow.hashCode());
    result = 31 * result + (this.lastColumn == null ? 0 : this.lastColumn.hashCode());
    result = 31 * result + (this.lastRow == null ? 0 : this.lastRow.hashCode());
    return result;
  }

  public boolean equals(Object obj) {
    if (this == obj) {
      return true;
    } else if (obj == null) {
      return false;
    } else if (this.getClass() != obj.getClass()) {
      return false;
    } else {
      ExcelMergedRegion other = (ExcelMergedRegion) obj;
      if (this.firstColumn == null) {
        if (other.firstColumn != null) {
          return false;
        }
      } else if (!this.firstColumn.equals(other.firstColumn)) {
        return false;
      }

      if (this.firstRow == null) {
        if (other.firstRow != null) {
          return false;
        }
      } else if (!this.firstRow.equals(other.firstRow)) {
        return false;
      }

      if (this.lastColumn == null) {
        if (other.lastColumn != null) {
          return false;
        }
      } else if (!this.lastColumn.equals(other.lastColumn)) {
        return false;
      }

      if (this.lastRow == null) {
        if (other.lastRow != null) {
          return false;
        }
      } else if (!this.lastRow.equals(other.lastRow)) {
        return false;
      }

      return true;
    }
  }

  public Boolean isMergedRegion(Integer row, Integer column) {
    return row >= this.firstRow
        && row <= this.lastRow
        && column >= this.firstColumn
        && column <= this.lastColumn;
  }
}
