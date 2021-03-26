using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VSKeyExtractor
{
    record Product(string Name, string GUID, string MPC);

    class Program
    {
        static readonly List<Product> Products = new()
        {
            new Product("Visual Studio 2015 Enterprise"       , "4D8CFBCB-2F6A-4AD2-BABF-10E28F6F2C8F", "07060"),
            new Product("Visual Studio 2015 Professional"     , "4D8CFBCB-2F6A-4AD2-BABF-10E28F6F2C8F", "07062"),

            new Product("Visual Studio 2017 Enterprise"       , "5C505A59-E312-4B89-9508-E162F8150517", "08860"),
            new Product("Visual Studio 2017 Professional"     , "5C505A59-E312-4B89-9508-E162F8150517", "08862"),
            new Product("Visual Studio 2017 Test Professional", "5C505A59-E312-4B89-9508-E162F8150517", "08866"),

            new Product("Visual Studio 2019 Enterprise"       , "41717607-F34E-432C-A138-A3CFD7E25CDA", "09260"),
            new Product("Visual Studio 2019 Professional"     , "41717607-F34E-432C-A138-A3CFD7E25CDA", "09262"),
        };

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
                var unicode = new UnicodeEncoding();
                var str = unicode.GetString(secret);
                foreach (var sub in str.Split('\0'))
                {
                    var match = Regex.Match(sub, @"\w{5}-\w{5}-\w{5}-\w{5}-\w{5}");
                    if (match.Success)
                    {
                        Console.WriteLine($"Found key for {product.Name}: {match.Captures[0]}");
                    }
                }
            }
            catch (Exception) { }
        }
    }
}

