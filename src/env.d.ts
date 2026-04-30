declare global {
  interface Env {
    SESSION_SECRET?: string;
    APP_USERS_JSON?: string;
    GOOGLE_SERVICE_ACCOUNT_JSON?: string;
  }
}

export {};
