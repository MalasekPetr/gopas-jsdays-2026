function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

export const TENANT_ID = requireEnv("TENANT_ID");
export const CLIENT_ID = requireEnv("CLIENT_ID");
export const CLIENT_SECRET = requireEnv("CLIENT_SECRET");
export const SHAREPOINT_SITE_URL = requireEnv("SHAREPOINT_SITE_URL");
export const TARGET_USER_EMAIL = requireEnv("TARGET_USER_EMAIL");
export const AUDIT_LIST_NAME = requireEnv("AUDIT_LIST_NAME");
