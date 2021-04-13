using System;
using Excel;
using IronXL;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace qual
{
    // backend engine for data parsing
    class Parser
    {
        Dictionary<String,Employee> employees;  // dict <ID,emp> of all employees
        Dictionary<String, Employee> empSaved;  // employees from save file
        List<Employee> expiring;                // list of expiring employees
        Dictionary<String, String> emailByID;   // dict <ID, email> of emails for employees
        string newReportSrc = null;
        string saveFileSrc = "QualReport.xlsx";

        // parse data from a given exported qualification sheet
        public void ParseSheet(String fileSource)
        {
            if (fileSource != null)
            {
                if (System.IO.File.Exists(saveFileSrc))
                {
                    newReportSrc = fileSource;
                }
                else
                {
                    // file not found
                    return;
                }
            }
            else
            {
                // no file given
                return;
            }
            worksheet ws = Workbook.Worksheets(fileSource).ElementAt(0);

            employees = new Dictionary<String, Employee>();
            expiring = new List<Employee>();
            Excel.Cell[] row;
            String id;
            Employee emp;
            int[] dmy;
            foreach (Row r in ws.Rows)
            {
                if (r == null) break;
                row = r.Cells;
                if (row[4] == null) continue; // bottom (usually) line is half empty
                id = row[0].Text;
                if (id == "ID") continue; // header row

                // get existing employee if exists
                if (employees.ContainsKey(id))
                {
                    emp = employees[id];
                }
                // othewise create a new one
                else
                {
                    emp = new Employee()
                    {
                        ID = id,
                        LastName = row[1].Text,
                        GivenName = row[2].Text
                    };
                    // add to dictionary
                    employees.Add(id, emp);
                }

                // parse the certification
                dmy = ExcelSerialDateToDMY(int.Parse(row[6].Value));
                Certification cert = new Certification()
                {
                    Name = row[3].Text,
                    Expiry = new DateTime(dmy[2], dmy[1], dmy[0])
                };
                cert.DaysLeft = (int)(cert.Expiry - DateTime.Now).TotalDays;

                // append certification to employee
                emp.Certifications[cert.Name] = cert;

                // mark expiring qual
                if (cert.DaysLeft <= 60)
                {
                    emp.Expiring = true;
                }
            }
            ParseSaved();
        }

        public void ParseSaved()
        {
            worksheet ws;
            if (System.IO.File.Exists(saveFileSrc))
            {
                // open sheet 1
                ws = Workbook.Worksheets(saveFileSrc).ElementAt(0);
            }
            else
            {
                // file not found
                return;
            }
            
            // repeat parsing for saved spreadsheet

            Excel.Cell[] row;
            String id;
            Employee emp;
            int[] dmy, emailed;
            foreach (Row r in ws.Rows)
            {
                if (r == null) break;
                row = r.Cells;
                if (row[4] == null) continue; // bottom (usually) line is half empty
                id = row[0].Text;
                if (id == "ID") continue; // header row

                // get existing employee if exists
                if (employees.ContainsKey(id))
                {
                    emp = employees[id];
                }
                // othewise nothing to update
                else
                {
                    continue;
                }

                // parse the certification
                dmy = ExcelSerialDateToDMY(int.Parse(row[6].Value));

                // if no emailed date, nothing to update
                if (row[7] == null) continue;
                // grab date
                emailed = ExcelSerialDateToDMY(int.Parse(row[7].Value));

                // new certification if updatable
                Certification cert = new Certification()
                {
                    Name = row[3].Text,
                    Expiry = new DateTime(dmy[2], dmy[1], dmy[0]),
                    EmailedDate = new DateTime(emailed[2], emailed[1], emailed[0]),
                    Emailed = true
                };
                cert.DaysLeft = (int)(cert.Expiry - DateTime.Now).TotalDays;


                // append certification to employee
                emp.Certifications.Add(cert.Name, cert);

                // mark expiring qual
                if (cert.DaysLeft <= 60)
                {
                    emp.Expiring = true;
                }
                Console.WriteLine(cert);
            }
        }

        public void SaveData()
        {
            using (StreamWriter writer = new StreamWriter(saveFileSrc, true))
            {
                WorkBook newReport = WorkBook.Load(newReportSrc);


            }
        }

        public void ParseEmails(worksheet ws)
        {
            emailByID = new Dictionary<string, string>();
            String id, email;
            foreach (Row r in ws.Rows)
            {
                id = r.Cells[0].Value; // Value to get number without formatting
                email = r.Cells[1].Text;
                emailByID.Add(id, email);
                
            }
        }

        private int[] ExcelSerialDateToDMY(int serialDate)
        {
            int[] dmy = new int[3];

            // 29-02-1900 bug backward compatibility
            if (serialDate == 60)
            {
                dmy[0] = 29;
                dmy[1] = 2;
                dmy[2] = 1900;

                return dmy;
            }
            else if (serialDate < 60)
            {
                serialDate++;
            }

            // Modified Julian to DMY calculation with an addition of 2415019
            int l = serialDate + 68569 + 2415019;
            int n = (int)((4 * l) / 146097);
            l = l - (int)((146097 * n + 3) / 4);
            int i = (int)((4000 * (l + 1)) / 1461001);
            l = l - (int)((1461 * i) / 4) + 31;
            int j = (int)((80 * l) / 2447);
            dmy[0] = l - (int)((2447 * j) / 80);
            l = (int)(j / 11);
            dmy[1] = j + 2 - (12 * l);
            dmy[2] = 100 * (n - 49) + i + l;

            return dmy;
        }

        // return list of all employees
        public List<Employee> GetEmployees()
        {
            return employees.Values.ToList();
        }

        // create and return a list of RowEntry , a formatted struct for the data grid
        public List<RowEntry> GetEntries()
        {
            var certs = new List<RowEntry>();
            foreach (Employee e in employees.Values)
            {
                foreach (Certification c in e.Certifications.Values)
                {
                    certs.Add(new RowEntry()
                    {
                        ID = e.ID,
                        LastName = e.LastName,
                        GivenName = e.GivenName,
                        Cert = c.Name,
                        Expiry = c.Expiry.ToString("d"),
                        DaysLeft = c.DaysLeft,
                        Emailed = c.Emailed,
                        // if no date, default to "None"
                        EmailedDate = c.Emailed ? c.EmailedDate.ToString("d") : "None"
                    });
                }
            }
            return certs;
        }
        public List<Employee> GetExpiring()
        {
            return expiring;
        }

        public Dictionary<String,String> GetEmails()
        {
            return emailByID;
        }

        // send an email to a given employee by row entry
        public void SendEmail(RowEntry row)
        {
            Employee emp = employees[row.ID];

            string name = emp.GivenName;
            string cert = row.Cert;
            string date = row.Expiry;

            string address = emailByID[emp.ID];
            string subject = "Expiring Qualification";
            string cc = "";
            string bcc = "";

            // newline character %0D%0A
            string body = String.Format("Hi {0},%0D%0A" +
                        "Your \'{1}\' is expiring in less than 60 days, on {2}.%0D%0A" +
                        "Please provide proof of certification when you receive it.%0D%0A%0D%0A" +
                        "Thanks,%0D%0AJenn", name, cert, date);

            // pop open a new email window
            Process.Start(String.Format("mailto:{0}?subject={1}&cc={2}&bcc={3}&body={4}",
                address, subject, cc, bcc, body));

            // update emailed status for employee's certification
            foreach (Certification c in emp.Certifications.Values)
            {
                if (c.Name == cert)
                {
                    c.Emailed = true;
                    c.EmailedDate = DateTime.Now;
                    break;
                }
            }
        }
    }
}
