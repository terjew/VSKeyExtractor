using Microsoft.Win32;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VSKeyExtractor {
  readonly struct Product {
    public readonly string Name { get; }
    public readonly string Edition { get; }
    public readonly string FullName { get; }

    public readonly string GUID { get; }
    public readonly string MPC { get; }
    public readonly string Path { get; }
    public readonly string ProductKey { get; }

    public Product(string Name, string Edition, string GUID, string MPC)
    {
      this.Name = Name;
      this.Edition = Edition;
      this.FullName = Name + " " + Edition;
      this.GUID = GUID;
      this.MPC = MPC;
      this.Path = $"HKEY_CLASSES_ROOT\\Licenses\\{GUID}\\{MPC}";
      this.ProductKey = RegUtil.TryGetLicense(RegUtil.TryGetValue(this.Path));
    }
  }

  public class Program {
    static readonly Product[] Products = new Product[]
    {
      new Product("Visual Studio 2012", "*"                , "77550D6B-6352-4E77-9DA3-537419DF564B", "05500"), // This is just a guess as to what the MPC number might be around

      new Product("Visual Studio 2013", "Professional"     , "E79B3F9C-6543-4897-BBA5-5BFB0A02BB5C", "06177"),

      new Product("Visual Studio 2015", "Enterprise"       , "4D8CFBCB-2F6A-4AD2-BABF-10E28F6F2C8F", "07060"),
      new Product("Visual Studio 2015", "Professional"     , "4D8CFBCB-2F6A-4AD2-BABF-10E28F6F2C8F", "07062"),

      new Product("Visual Studio 2017", "Enterprise"       , "5C505A59-E312-4B89-9508-E162F8150517", "08860"),
      new Product("Visual Studio 2017", "Professional"     , "5C505A59-E312-4B89-9508-E162F8150517", "08862"),
      new Product("Visual Studio 2017", "Test Professional", "5C505A59-E312-4B89-9508-E162F8150517", "08866"),

      new Product("Visual Studio 2019", "Enterprise"       , "41717607-F34E-432C-A138-A3CFD7E25CDA", "09260"),
      new Product("Visual Studio 2019", "Professional"     , "41717607-F34E-432C-A138-A3CFD7E25CDA", "09262"),

      
      new Product("Visual Studio 2022", "Enterprise"       , "1299B4B9-DFCC-476D-98F0-F65A2B46C96D", "09660"),
      new Product("Visual Studio 2022", "Professional"     , "1299B4B9-DFCC-476D-98F0-F65A2B46C96D", "09662"),
      new Product("Visual Studio 2022", "Prerelease Trial" , "B16F0CF0-8AD1-4A5B-87BC-CB0DBE9C48FC", "09562"),
    };


    static void Main()
    {
      string DIVIDER = new String('-', 79) + '\n';
      int padding = 0;
      foreach (Product product in Products) { padding = Math.Max(product.Edition.Length, padding); }

      List<string> knownProductKeys = new List<string>();

      foreach (var group in Products.Where(x=>!string.IsNullOrWhiteSpace(x.ProductKey)).GroupBy(x=>x.Name)) {
        Console.WriteLine(DIVIDER + group.Key);
        foreach (var product in group.ToArray()) {
          knownProductKeys.Add(product.ProductKey);
          Console.WriteLine($"  {(product.Edition  + ": ").PadRight(padding+2)}{product.ProductKey}");
        }
      }

      using var rkLicenses = Registry.ClassesRoot.TryOpenKey("Licenses");
      foreach (var subkeyGuid in rkLicenses.GetSubKeys(false)) {
        foreach (var subkeyMPC in subkeyGuid.GetSubKeys(false)) {
          string productKey = RegUtil.TryGetLicense(subkeyMPC.TryGetValue());
          if (productKey != null) {
            if (knownProductKeys.Contains(productKey)) { continue; }

            Console.WriteLine('\n' + DIVIDER + $"UNKNOWN PRODUCTKEY:  {productKey} \n  found at: {subkeyMPC.Name}");

            string possibleVersion = Products.Where(x => x.Path.StartsWith(subkeyGuid.Name)).FirstOrDefault().Name; // ?? Products.Where(x => x.Path.StartsWith(subkeyGuid.Name)).FirstOrDefault().Name;
            if (!string.IsNullOrWhiteSpace(possibleVersion)) {
              Console.WriteLine($"  Possibly associated with {possibleVersion}");
            } else {
              Console.WriteLine("  Possibly belonging to a version earlier than 2012 or a pre-release trial. \n  Closest guess is a " + GetAssociatedVersionGuessMessage(subkeyMPC.Name));
            }
          }
        }
      }

      Console.ReadKey();
    }

    private static string GetAssociatedVersionGuessMessage(string path)
    {
      var parts = path.Split('\\', '/');
      Array.Reverse(parts);
      var mpcVer = TryConvertHex(parts[0]);
      if (mpcVer == 0) { return null; }


      var sortedByEdition = Products.Select(x => { return new KeyValuePair<UInt64, Product>(TryConvertHex(x.MPC), x); }).OrderBy(kvp => kvp.Key).ToArray();

      if (sortedByEdition[0].Key > mpcVer) { return $"version earlier than {sortedByEdition[0].Value.FullName}"; }
      if (sortedByEdition[0].Key == mpcVer) { return sortedByEdition[0].Value.FullName; }
      if (sortedByEdition[sortedByEdition.Length-1].Key < mpcVer) { return $"version later than {sortedByEdition[sortedByEdition.Length-1].Value.FullName}"; }
      if (sortedByEdition[sortedByEdition.Length-1].Key == mpcVer) { return sortedByEdition[sortedByEdition.Length-1].Value.FullName; }

      for (var i = 1; i < sortedByEdition.Length; i++) {
        var kvpPrev = sortedByEdition[i-1];
        var kvpCurrent = sortedByEdition[i];
        if (kvpPrev.Key <= mpcVer) {
          if (kvpCurrent.Key >= mpcVer) {
            return $"version between {kvpPrev.Value.FullName} and {kvpCurrent.Value.FullName}";
          }
        }
      }
      return null;

      static UInt64 TryConvertHex(string hexString)
      {
        try { return Convert.ToUInt64(hexString, 16); } catch (Exception ex) { return 0; }
      }
    }

  }

  internal static class RegUtil {
    static readonly UnicodeEncoding Unicode = new UnicodeEncoding();
    static readonly Regex RGX = new Regex(@"(\w{5}-\w{5}-\w{5}-\w{5}-\w{5})");

    public static string TryGetLicense(object value) => TryGetLicense(TryDecrypt(value as byte[]));

    public static string TryGetLicense(string decrypted)
    {
      if (string.IsNullOrWhiteSpace(decrypted)) { return null; }
      var match = RGX.Match(decrypted);
      return match.Success ? match.Groups[1].Value : null;
    }

    public static string TryDecrypt(byte[] bytes)
    {
      if (bytes == null) { return null; }
      try { return Unicode.GetString(ProtectedData.Unprotect(bytes, null, DataProtectionScope.CurrentUser)); } catch (Exception) { return null; }
    }

    public static object TryGetValue(string path, string valueName = "", object defaultValue = null)
    {
      try { return Registry.GetValue(path, valueName, defaultValue); } catch (Exception ex) { return defaultValue; }
    }

    public static object TryGetValue(this RegistryKey rk, string valueName = "", object defaultValue = null)
    {
      try { return rk.GetValue(valueName, defaultValue); } catch (Exception ex) { return defaultValue; }
    }

    public static IEnumerable<RegistryKey> GetSubKeys(this RegistryKey regKey, bool writable = false)
    {
      string[] subkeyNames = regKey.TryGetSubKeyNames();
      for (int i = 0; i < subkeyNames.Length; i++) {
        RegistryKey subkey;
        try { subkey = regKey.OpenSubKey(subkeyNames[i], writable); } catch (Exception ex) { continue; }
        yield return subkey;
        subkey.Dispose();
      }
    }


    public static string[] TryGetSubKeyNames(this RegistryKey rk)
    {
      try { return rk.GetSubKeyNames(); } catch (Exception ex) { return new string[0]; }
    }


    public static RegistryKey TryOpenKey(this RegistryKey rk, string name, bool writable = false)
    {
      try { return rk.OpenSubKey(name, writable); } catch (Exception ex) { return null; }
    }
  }
}

