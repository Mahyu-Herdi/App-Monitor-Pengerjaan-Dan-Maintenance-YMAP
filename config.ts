export const APP_CONFIG = {
  password: "M24yu@mhy"
};

if (typeof window !== 'undefined') {
  (window as any).APP_CONFIG = APP_CONFIG;
}
