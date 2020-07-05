using Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

public class Employee
{
    /* Represents one employee, with all their certifications in a list */
    public String ID { get; set; }
    public string LastName { get; set; }
    public string GivenName { get; set; }
    // whether employee has expiring certification
    public bool Expiring { get; set; } = false;
    public Dictionary<String,Certification> Certifications { get; set; } = new Dictionary<String,Certification>();
}

public class Certification
{
    /* Represents one certification */
    public string Name { get; set; }
    public DateTime Expiry { get; set; }
    public int DaysLeft { get; set; }
    public bool Emailed { get; set; } = false;
    public DateTime EmailedDate { get; set; }

}
public class RowEntry
{
    /* Combined form of Employee and Certification
     * Represents one row in the display table (takes all values from the struct)
     * One row per certification per person
     * Dates are Strings so they can be pre-formatted
     */
    public String ID { get; set; }
    public string LastName { get; set; }
    public string GivenName { get; set; }
    public String Cert { get; set; }
    public String Expiry { get; set; }
    public int DaysLeft { get; set; }
    public bool Expiring { get; set; } = false;
    public bool Emailed { get; set; } = false;
    public String EmailedDate { get; set; }
}