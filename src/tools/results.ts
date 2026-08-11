type JsonObject = { [key: string]: unknown };

export function jsonResult(output: unknown) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(output, null, 2) }],
    structuredContent: output as JsonObject
  };
}

export function formatError(error: unknown): JsonObject {
  if (error instanceof Error) {
    const details: JsonObject = {
      name: error.name,
      message: error.message
    };
    if ('code' in error) details.code = (error as { code?: unknown }).code;
    if ('status' in error) details.status = (error as { status?: unknown }).status;
    if ('responseBody' in error) details.responseBody = (error as { responseBody?: unknown }).responseBody;
    return details;
  }
  return { message: String(error) };
}

export function describeTool(summary: string, chineseIntents: string[]): string {
  return summary + ' Chinese intents / 中文意图: ' + chineseIntents.join(', ') + '.';
}

export async function runTool(handler: () => Promise<unknown>) {
  try {
    return jsonResult(await handler());
  } catch (error) {
    const output = { error: formatError(error) };
    return {
      content: [{ type: 'text' as const, text: JSON.stringify(output, null, 2) }],
      structuredContent: output,
      isError: true
    };
  }
}

export function errorResult(code: string, message: string) {
  const output = { error: { code, message } };
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(output, null, 2) }],
    structuredContent: output,
    isError: true
  };
}
