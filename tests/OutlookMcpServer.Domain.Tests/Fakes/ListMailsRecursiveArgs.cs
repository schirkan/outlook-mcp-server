using OutlookMcpServer.Domain.Models.Common;

namespace OutlookMcpServer.Domain.Tests.Fakes;

/// <summary>
/// Test-Argument-Container fuer <see cref="FakeInteropAdapter.ListMailsRecursiveAsync"/>.
/// </summary>
public sealed record ListMailsRecursiveArgs(
    IReadOnlyList<string> Scope,
    int Top,
    string? Filter,
    BodyFormat BodyFormat);
