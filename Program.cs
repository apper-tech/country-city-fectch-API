using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CountryCityCrowler
{
    class Program
    {
        public const string ConnectionString = "Data Source=Lenovo-IdeaPad;Initial Catalog=DynamicEservice-RealEstate;User ID=sa;Password=admini;Connect Timeout=15;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultipleActiveResultSets=true;MultiSubnetFailover=False;";
        public static List<RegionInfo> ls = new List<RegionInfo>();
        static void Main(string[] args)
        {
            ls = GetCountryList();
            ls = ls.OrderBy(o => o.TwoLetterISORegionName).ToList<RegionInfo>();
            string CityPath = "../../Files/CitiesWithAlternatives.xlsx";
            DataTable dt = ReadDataExcel(CityPath);
            CrawlData(dt, 7505, 9999);
            Console.Write("Press Any Key to Exit");
            Console.ReadLine();
        }
        public static void CrawlData(DataTable dt,int rowstart,int rowend)
        {
            List<string> res = new List<string>();
            List<City> list = new List<City>();
            int count = 1;
            foreach (DataRow item in dt.Rows)
            {
                if (count >= rowstart && count <=rowend)
                {
                   // Console.Clear();
                    Console.WriteLine("Downloading.....");
                    string engname = item[0] + ""; string alt = item[1] + ""; string lat = item[2] + ""; string lng = item[3] + ""; string reg = item[4] + "";
                    string lang = GetLang(reg);
                    string natname = GetGeocodingData(lat, lng, reg, engname, lang);
                    if (natname == "OVER_QUERY_LIMIT")
                    {
                        Console.WriteLine("Out of Quata!!!!");
                        break;
                    }
                    if (natname == "ZERO_RESULTS")
                    {
                        Console.WriteLine("Not Found!!!!");
                        continue;
                    }
                    string natinlatinname = GetLocalNativeLatinName(alt);
                    string localres = string.Format("Address : {0} , Name: {1} ,Latin: {3}, Lang: {2}", engname, natname, lang, natinlatinname);
                    Console.WriteLine(localres);
                    Console.WriteLine(res.Count + "/" + dt.Rows.Count + "|" + (DateTime.Now - System.Diagnostics.Process.GetCurrentProcess().StartTime).Minutes + ":" + (DateTime.Now - System.Diagnostics.Process.GetCurrentProcess().StartTime).Seconds + " Elapsed");
                    //Console.WriteLine(res.Count);
                   // Console.ReadLine();
                    res.Add(localres);
                    list.Add(new City { Id=count.ToString(),Code = reg, Name = engname, Native = natname, Latin = natinlatinname,Lat=lat,Lng=lng });
                }
                count++;
            }
            File.WriteAllLines("text.txt", res);
            foreach (var item in list)
            {
                 InsertCity(item.Code, item.Name, item.Native, item.Latin,item.Lat,item.Lng);
                //UpdateCityLatLng(item.Lat, item.Lng, item.Id);
            }
        }
        
        public static DataTable ReadDataExcel(string filepath)
        {
            FileStream stream = File.Open(filepath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();
            DataTable dt = new DataTable();
            dt = result.Tables[0];
            return dt;
        }
        public static List<RegionInfo> GetCountryList()
        {

            List<RegionInfo> objCountries = new List<RegionInfo>();
            foreach (CultureInfo objCultureInfo in CultureInfo.GetCultures(CultureTypes.SpecificCultures))
            {
                RegionInfo objRegionInfo = new RegionInfo(objCultureInfo.Name);
                if (!CheckReigonKey(objCountries, objRegionInfo.GeoId))
                    objCountries.Add(objRegionInfo);
            }
            return objCountries;
        }
        private static bool CheckReigonKey(List<RegionInfo> ls, int key)
        {
            foreach (var item in ls)
            {
                if (item.GeoId == key)
                    return true;
            }
            return false;
        }
        private static void WriteTabTextFile(string cell1, string cell2, string cell3, string filename)
        {
            string line = string.Format("{0}	{1} 	{2}", cell1, cell2, cell3);
            FileStream fs = new FileStream(filename, FileMode.Append, FileAccess.Write);
            StreamWriter file = new StreamWriter(fs);
            file.WriteLine(line);
            file.Close();
        }
        private static void SetQuery(string sql)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
            }
        }
        public static void InsertCountry(string code, string name, string native)
        {
            string sql = @"INSERT INTO Countries (Country_Code,Country_Name,Country_Native_Name) values ('" + code + "','" + name + "',N'" + native + "')";
            SetQuery(sql);
        }
        public static void InsertCity(string code, string name, string native, string latin,string lat,string lng)
        {
            string sql = @"Insert Into Cities (City_Name,City_Native_Name,City_Latin_Name,Country_ID,Latitude,Longitude) values ('" + name + "',N'" + native + "','" + latin + "',(select Countries.Country_ID from Countries where Countries.Country_Code='" + code + "'),'"+lat+"','"+lng+"');";
            SetQuery(sql);
        }
        public static void UpdateCityLatLng(string lat, string lng, string id)
        {
            string sql = "update Cities set Latitude='" + lat + "',Longitude='" + lng + "' where Cities.City_ID=" + id + ";";
            SetQuery(sql);
        }
        public static string GetGeocodingData(string latitude, string longitude, string reigon, string address, string language)
        {
            string url = string.Format("http://maps.googleapis.com/maps/api/geocode/xml?latlng={0},{1}&sensor=false&language={2}&address={3}&reigon={4}", latitude, longitude, language, address, reigon);
            XElement xml = XElement.Load(url);
            if (xml.Element("status").Value == "OK")
            {
                string res = string.Format("{0}", xml.Element("result").Element("address_component").Element("long_name").Value);
                return res;
            }
            if(xml.Element("status").Value== "ZERO_RESULTS")
            return "ZERO_RESULTS";
            return xml.Element("status").Value;
        }
        public static string GetLang(string countryCode)
        {
            foreach (RegionInfo item in ls)
            {
                if (item.TwoLetterISORegionName == countryCode)
                {
                    return item.Name;
                }

            }
            return "en";
        }
        public static string GetLocalNativeLatinName(string commaSep)
        {
            string[] result = commaSep.Split(',');
            if (result.Length > 12)
                return result[12];
            return result[0];
        }
    }
    public class City
    {
        string id;
        string code;
        string name;
        string native;
        string latin;
        string lat;
        string lng;

        public string Code
        {
            get
            {
                return code;
            }

            set
            {
                code = value;
            }
        }

        public string Name
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
            }
        }

        public string Native
        {
            get
            {
                return native;
            }

            set
            {
                native = value;
            }
        }

        public string Latin
        {
            get
            {
                return latin;
            }

            set
            {
                latin = value;
            }
        }

        public string Lat
        {
            get
            {
                return lat;
            }

            set
            {
                lat = value;
            }
        }

        public string Lng
        {
            get
            {
                return lng;
            }

            set
            {
                lng = value;
            }
        }

        public string Id
        {
            get
            {
                return id;
            }

            set
            {
                id = value;
            }
        }
    }
}
