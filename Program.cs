using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VSKeyExtractor;

/// <summary>
/// Console entry point for extracting Visual Studio product license keys from the Windows registry.
/// </summary>
class Program
{
    /// <summary>
    /// Known Visual Studio product names, registry product codes, and license identifiers.
    /// The application iterates over this list and tries to locate a matching encrypted license entry
    /// under the registry path used by Visual Studio installations.
    /// </summary>
    private static readonly IReadOnlyList<Product> Products =
    [
        // Visual Studio 2012 family.
        new Product("Visual Studio Express 2012 for Windows Phone"  , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04937"),
        new Product("Visual Studio Professional 2012"               , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04938"),
        new Product("Visual Studio Ultimate 2012"                   , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04940"),
        new Product("Visual Studio Premium 2012"                    , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04941"),
        new Product("Visual Studio Test Professional 2012"          , new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "04942"),
        new Product("Visual Studio Express 2012 for Windows Desktop", new Guid("77550D6B-6352-4E77-9DA3-537419DF564B"), "05695"),

        // Visual Studio 2013 family.
        new Product("Visual Studio 2013 Professional"               , new Guid("E79B3F9C-6543-4897-BBA5-5BFB0A02BB5C"), "06177"),
        new Product("Visual Studio 2013 Ultimate"                   , new Guid("E79B3F9C-6543-4897-BBA5-5BFB0A02BB5C"), "06181"),

        // Visual Studio 2015 family.
        new Product("Visual Studio 2015 Enterprise"                 , new Guid("4D8CFBCB-2F6A-4AD2-BABF-10E28F6F2C8F"), "07060"),
        new Product("Visual Studio 2015 Professional"               , new Guid("4D8CFBCB-2F6A-4AD2-BABF-10E28F6F2C8F"), "07062"),

        // Visual Studio 2017 family.
        new Product("Visual Studio 2017 Enterprise"                 , new Guid("5C505A59-E312-4B89-9508-E162F8150517"), "08860"),
        new Product("Visual Studio 2017 Professional"               , new Guid("5C505A59-E312-4B89-9508-E162F8150517"), "08862"),
        new Product("Visual Studio 2017 Test Professional"          , new Guid("5C505A59-E312-4B89-9508-E162F8150517"), "08866"),

        // Visual Studio 2019 family.
        new Product("Visual Studio 2019 Enterprise"                 , new Guid("41717607-F34E-432C-A138-A3CFD7E25CDA"), "09260"),
        new Product("Visual Studio 2019 Professional"               , new Guid("41717607-F34E-432C-A138-A3CFD7E25CDA"), "09262"),

        // Visual Studio 2022 family.
        new Product("Visual Studio 2022 Enterprise"                 , new Guid("1299B4B9-DFCC-476D-98F0-F65A2B46C96D"), "09660"),
        new Product("Visual Studio 2022 Professional"               , new Guid("1299B4B9-DFCC-476D-98F0-F65A2B46C96D"), "09662"),

        // Visual Studio 2026 Insider builds.
        new Product("Visual Studio 2026 Enterprise Insider"         , new Guid("0EB1B2EC-090C-4540-B219-F529C658360C"), "09760"),
        new Product("Visual Studio 2026 Professional Insider"       , new Guid("0EB1B2EC-090C-4540-B219-F529C658360C"), "09762"),

        // Visual Studio 2026 release builds.
        new Product("Visual Studio 2026 Enterprise"                 , new Guid("97372B8F-5B80-4DA7-8476-FF55D6368CBD"), "09860"),
        new Product("Visual Studio 2026 Professional"               , new Guid("97372B8F-5B80-4DA7-8476-FF55D6368CBD"), "09862"),

    ];

    /// <summary>
    /// Starts the extraction process for every known product in the static product list.
    /// </summary>
    static void Main()
    {
        foreach (var product in Products) ExtractLicense(product);
    }

    /// <summary>
    /// Reads the encrypted license blob for a specific product from the registry, decrypts it,
    /// and searches the decrypted text for a key formatted as five groups of five alphanumeric
    /// characters separated by hyphens.
    /// </summary>
    /// <param name="product">The product metadata used to locate the registry entry.</param>
    private static void ExtractLicense(Product product)
    {
        // Visual Studio stores encrypted license data beneath this registry path.
        var encrypted = Registry.GetValue($"HKEY_CLASSES_ROOT\\Licenses\\{product.GUID}\\{product.MPC}", "", null);
        if (encrypted == null) return;
        try
        {
            // Decrypt the protected blob using the current user's DPAPI context.
            var secret = ProtectedData.Unprotect((byte[])encrypted, null, DataProtectionScope.CurrentUser);
            // Convert the decrypted binary data into a Unicode string for token scanning.
            var str = Encoding.Unicode.GetString(secret);
            foreach (var sub in str.Split('\0'))
                if (!string.IsNullOrWhiteSpace(sub))
                {
                    Debug.WriteLine($"sub: {sub}");
                    // Search each token for a license-key-shaped value.
                    var match = Regex.Match(sub, @"\w{5}-\w{5}-\w{5}-\w{5}-\w{5}");
                    if (match.Success)
                    {
                        Console.WriteLine($"Found key for {product.Name}: {match.Captures[0]}");
                    }
                }
        }
        catch (Exception)
        {
            // Ignore malformed or inaccessible registry entries and continue with the next product.
        }
    }
}
