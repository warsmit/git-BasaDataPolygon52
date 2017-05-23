using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization.Json;
using System.IO;
using System.Runtime.Serialization;


namespace BasaDataForPolygon52
{
    class JsonJob
    {
        public void WriteInDailyReport(string day)
        {
            DailyReport library;
            using (Stream stream = File.OpenRead(day + ".json"))
            {
                var serializer = new DataContractJsonSerializer(typeof(DailyReport));
                library = (DailyReport)serializer.ReadObject(stream);
            }

            using (Stream stream = File.OpenWrite("test.json"))
            {
                var serializer = new DataContractJsonSerializer(typeof(DailyReport));

                library.Receipt.Add(new Receipt()
                {
                    VendorCode = "ak",
                    Name = "AK47",
                    Count = "1",
                    Price = "2000",
                });

                serializer.WriteObject(stream, library);
            }
        }

        public void ReadDailyReport(string day)
        {
            using (Stream stream = File.OpenRead(day + ".json"))
            {
                var serializer = new DataContractJsonSerializer(typeof(DailyReport));
                DailyReport library = (DailyReport)serializer.ReadObject(stream);
            }
        }
    }

    [DataContract]
    class Receipt
    {
        [DataMember(Name = "vendorCode")]
        public string VendorCode { get; set; }
        [DataMember(Name = "name")]
        public string Name { get; set; }
        [DataMember(Name = "count")]
        public string Count { get; set; }
        [DataMember(Name = "price")]
        public string Price { get; set; }
    }

    [DataContract]
    class DailyReport
    {
        [DataMember(Name = "receipt")]
        public List<Receipt> Receipt { get; set; }
        /*
        [DataMember(Name = "src")]
        public string Src { get; set; }
        [DataMember(Name = "id")]
        public string Id { get; set; }
        */
    }
}
