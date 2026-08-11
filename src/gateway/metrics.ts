type ToolMetric = {
  calls: number;
  errors: number;
  durationMs: number;
};

function escapeLabel(value: string): string {
  return value.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n');
}

export class GatewayMetrics {
  private readonly startedAt = Date.now();
  private readonly tools = new Map<string, ToolMetric>();

  record(toolName: string, category: string, success: boolean, durationMs: number): void {
    const key = `${toolName}\u0000${category}`;
    const metric = this.tools.get(key) ?? { calls: 0, errors: 0, durationMs: 0 };
    metric.calls += 1;
    metric.errors += success ? 0 : 1;
    metric.durationMs += durationMs;
    this.tools.set(key, metric);
  }

  snapshot(): object {
    return {
      uptimeSeconds: Math.floor((Date.now() - this.startedAt) / 1000),
      tools: [...this.tools.entries()].map(([key, metric]) => {
        const [toolName, category] = key.split('\u0000');
        return { toolName, category, ...metric };
      })
    };
  }

  prometheus(): string {
    const lines = [
      '# HELP mcp_gateway_uptime_seconds Gateway process uptime in seconds.',
      '# TYPE mcp_gateway_uptime_seconds gauge',
      `mcp_gateway_uptime_seconds ${Math.floor((Date.now() - this.startedAt) / 1000)}`,
      '# HELP mcp_tool_calls_total Total MCP tool calls.',
      '# TYPE mcp_tool_calls_total counter',
      '# HELP mcp_tool_errors_total Total failed MCP tool calls.',
      '# TYPE mcp_tool_errors_total counter',
      '# HELP mcp_tool_duration_milliseconds_total Cumulative MCP tool execution time.',
      '# TYPE mcp_tool_duration_milliseconds_total counter'
    ];
    for (const [key, metric] of [...this.tools.entries()].sort(([left], [right]) => left.localeCompare(right))) {
      const [toolName, category] = key.split('\u0000');
      const labels = `tool="${escapeLabel(toolName)}",category="${escapeLabel(category)}"`;
      lines.push(`mcp_tool_calls_total{${labels}} ${metric.calls}`);
      lines.push(`mcp_tool_errors_total{${labels}} ${metric.errors}`);
      lines.push(`mcp_tool_duration_milliseconds_total{${labels}} ${metric.durationMs}`);
    }
    return `${lines.join('\n')}\n`;
  }

  resetForTests(): void {
    this.tools.clear();
  }
}

export const gatewayMetrics = new GatewayMetrics();
