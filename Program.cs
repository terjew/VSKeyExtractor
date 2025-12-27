using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VSKeyExtractor
{
    internal struct Product
    {
        public string Name;
        public Guid GUID;
        public string MPC;

        public Product(string item1, Guid item2, string item3)
        {
            Name = item1;
            GUID = item2;
            MPC = item3;
        }
    }

    class Program
    {
        private static readonly List<Product> products = new List<Product>()
        {
            new Product("Visual Studio Express 2012 for Windows Phone"  , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04937"),
            new Product("Visual Studio Professional 2012"               , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04938"),
            new Product("Visual Studio Ultimate 2012"                   , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04940"),
            new Product("Visual Studio Premium 2012"                    , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04941"),
            new Product("Visual Studio Test Professional 2012"          , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04942"),
            new Product("Visual Studio Express 2012 for Windows Desktop", new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "05695"),

            new Product("Visual Studio 2013 Professional"               , new Guid("E79B3F9C-6543-4897-BBA5-5BFB0A02BB5C"), "06177"),
            new Product("Visual Studio 2013 Ultimate"                   , new Guid("E79B3F9C-6543-4897-BBA5-5BFB0A02BB5C"), "06181"),

            new Product("Visual Studio 2015 Enterprise"                 , new Guid("4D8CFBCB-2F6A-4AD2-BABF-10E28F6F2C8F"), "07060"),
            new Product("Visual Studio 2015 Professional"               , new Guid("4D8CFBCB-2F6A-4AD2-BABF-10E28F6F2C8F"), "07062"),

            new Product("Visual Studio 2017 Enterprise"                 , new Guid("5C505A59-E312-4B89-9508-E162F8150517"), "08860"),
            new Product("Visual Studio 2017 Professional"               , new Guid("5C505A59-E312-4B89-9508-E162F8150517"), "08862"),
            new Product("Visual Studio 2017 Test Professional"          , new Guid("5C505A59-E312-4B89-9508-E162F8150517"), "08866"),

            new Product("Visual Studio 2019 Enterprise"                 , new Guid("41717607-F34E-432C-A138-A3CFD7E25CDA"), "09260"),
            new Product("Visual Studio 2019 Professional"               , new Guid("41717607-F34E-432C-A138-A3CFD7E25CDA"), "09262"),

            new Product("Visual Studio 2022 Enterprise"                 , new Guid("1299B4B9-DFCC-476D-98F0-F65A2B46C96D"), "09660"),
            new Product("Visual Studio 2022 Professional"               , new Guid("1299B4B9-DFCC-476D-98F0-F65A2B46C96D"), "09662"),

        };
        static readonly List<Product> Products = products;

        static void Main()
        {
            foreach (var product in Products) ExtractLicense(product);
        }

        private static void ExtractLicense(Product product)
        {
            var encrypted = Registry.GetValue($"HKEY_CLASSES_ROOT\\Licenses\\{product.GUID}\\{product.MPC}", "", null);
            if (encrypted == null) return;
            try
            {
                var secret = ProtectedData.Unprotect((byte[])encrypted, null, DataProtectionScope.CurrentUser);
                var str = Encoding.Unicode.GetString(secret);
                foreach (var sub in str.Split('\0'))
                    if (!string.IsNullOrWhiteSpace(sub))
                    {
                        Debug.WriteLine($"sub: {sub}");
                        var match = Regex.Match(sub, @"\w{5}-\w{5}-\w{5}-\w{5}-\w{5}");
                        if (match.Success)
                        {
                            Console.WriteLine($"Found key for {product.Name}: {match.Captures[0]}");
                        }
                    }
            }
            catch (Exception) {/*just void em*/}
        }
    }
}
