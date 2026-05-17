using System;

namespace VSKeyExtractor;

/// <summary>
/// Represents a product entry with its display name, GUID, and MPC value.
/// </summary>
/// <param name="Name"> Gets the product name. </param>
/// <param name="GUID"> Gets the product GUID. </param>
/// <param name="MPC"> Gets the product MPC value. </param>
internal record struct Product(string Name, Guid GUID, string MPC)
{
    /// <summary>
    /// Gets the product name.
    /// </summary>
    public string Name { get; } = Name;

    /// <summary>
    /// Gets the product GUID.
    /// </summary>
    public Guid GUID { get; } = GUID;

    /// <summary>
    /// Gets the product MPC value.
    /// </summary>
    public string MPC { get; } = MPC;
}
