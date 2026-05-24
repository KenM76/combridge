namespace ComBridge.Core;

/// <summary>
/// Description of one running COM-server instance discovered in the ROT.
/// Built by the host from the index + the plugin's <c>DescribeInstance</c> hook.
/// </summary>
public sealed record SessionInfo(int Index, int? Pid, string? Title, string Description);
