using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace ZipCodeConverter.Models
{
    [DataContract]
    public class Address
    {
        public Address() { }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="columns"></param>
        public Address(string[] columns)
        {
            this.ZipCode = columns[2].Substring(0, 3) + "-" + columns[2].Substring(3, 4);
            this.AddressItems = new List<string>();

            this.AddressItems.Add(columns[6]);
            this.AddressItems.Add(columns[7]);
            this.AddressItems.Add(columns[8].Split('(').First().Split('（').First().Replace("以下に掲載がない場合", ""));

            this.AddressItems.Add(columns[3]);
            this.AddressItems.Add(columns[4]);
            this.AddressItems.Add(columns[5]);
        }

        [DataMember(Order = 0,Name="value")]
        public string ZipCode { get; set; }

        [DataMember(Order = 1,Name="address")]
        public List<string> AddressItems{get;set;}
    }
}
