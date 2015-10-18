namespace XmlReporter
{
    using EntryPoint;
    using System;
    using System.Linq;
    using System.Text;
    using System.Xml;

    class Startup
    {
        private const string OutputFilePath = @"../../../Output/";

        static void Main(string[] args)
        {
            using (var db = new TelerikAcademyEntities())
            {
                var employees = db.Employees.AsQueryable().Take(10);

                Encoding encoding = Encoding.GetEncoding("windows-1251");

                using (XmlTextWriter writer = new XmlTextWriter(OutputFilePath, encoding))
                {
                    writer.Formatting = Formatting.Indented;
                    writer.IndentChar = '\t';
                    writer.Indentation = 1;

                    writer.WriteStartDocument();

                    writer.WriteStartElement("employess");
                    foreach (var employee in employees)
                    {
                        WriteXElement("employee", writer, db, employee);
                    }
                    writer.WriteEndElement();

                    writer.WriteEndDocument();
                }

                Console.WriteLine("Document {0} created.", OutputFilePath);
            }
        }

        private static void WriteXElement(string elementName, XmlWriter writer, TelerikAcademyEntities db, Object obj)
        {
            writer.WriteStartElement(elementName);

            var properties = obj.GetType().GetProperties()
                .Where(p => !p.GetGetMethod().IsVirtual);

            foreach (var property in properties)
            {
                var propertyValue = property.GetValue(obj);
                if (propertyValue != null)
                {
                    var nameLowerCase = property.Name.ToLower();
                    if (nameLowerCase == "departmentid" && !obj.GetType().Name.StartsWith(typeof(Department).Name))
                    {
                        var department = db.Departments
                            .Where(d => d.DepartmentID == (int)propertyValue)
                            .FirstOrDefault();
                        WriteXElement("department", writer, db, department);
                    }
                    else if (nameLowerCase == "addressid" && !obj.GetType().Name.StartsWith(typeof(Address).Name))
                    {
                        var address = db.Addresses
                            .Where(a => a.AddressID == (int)propertyValue)
                            .FirstOrDefault();
                        WriteXElement("address", writer, db, address);
                    }
                    else if (nameLowerCase == "townid" && !obj.GetType().Name.StartsWith(typeof(Town).Name))
                    {
                        var town = db.Towns
                            .Where(t => t.TownID == (int)propertyValue)
                            .FirstOrDefault();
                        WriteXElement("town", writer, db, town);
                    }
                    else
                    {
                        writer.WriteElementString(nameLowerCase, propertyValue.ToString());
                    }
                }
            }
            writer.WriteEndElement();
        }
    }
}