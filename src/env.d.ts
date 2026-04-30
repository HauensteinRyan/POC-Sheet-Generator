declare global {
  interface Env {
    AUTH_SERVICE: Fetcher;
    AUTH_PUBLIC_BASE?: string;
    AUTH_APP_KEY?: string;
    GOOGLE_SERVICE_ACCOUNT_JSON?: string;
  }
}

export {};
