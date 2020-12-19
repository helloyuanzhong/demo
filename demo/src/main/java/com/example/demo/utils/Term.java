package com.example.demo.utils;

import java.io.Serializable;
import java.util.Date;

public class Term implements Serializable {
  private Date startDate;
  private Date endDate;

  public Term(Date startDate, Date endDate) {
    if (startDate.after(endDate)) {

    } else {
      this.startDate = startDate;
      this.endDate = endDate;
    }
  }

  public Date getStartDate() {
    return this.startDate;
  }

  public void setStartDate(Date startDate) {
    this.startDate = startDate;
  }

  public Date getEndDate() {
    return this.endDate;
  }

  public void setEndDate(Date endDate) {
    this.endDate = endDate;
  }

  public Boolean contain(Date date) {
    if (date.equals(this.startDate)) {
      return true;
    } else {
      return !date.before(this.startDate) && !date.after(this.endDate) && !date.equals(this.endDate)
          ? true
          : false;
    }
  }

  @Override
  public int hashCode() {
    boolean prime = true;
    int result = 1;
    result = 31 * result + (this.endDate == null ? 0 : this.endDate.hashCode());
    result = 31 * result + (this.startDate == null ? 0 : this.startDate.hashCode());
    return result;
  }

  @Override
  public boolean equals(Object obj) {
    if (this == obj) {
      return true;
    } else if (obj == null) {
      return false;
    } else if (this.getClass() != obj.getClass()) {
      return false;
    } else {
      Term other = (Term) obj;
      if (this.endDate == null) {
        if (other.endDate != null) {
          return false;
        }
      } else if (!this.endDate.equals(other.endDate)) {
        return false;
      }

      if (this.startDate == null) {
        if (other.startDate != null) {
          return false;
        }
      } else if (!this.startDate.equals(other.startDate)) {
        return false;
      }

      return true;
    }
  }
}
