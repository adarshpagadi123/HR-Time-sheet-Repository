import type { WalnutContext } from './walnut';

/** @walnut_method
 * name: Login to Application
 * description: Login with ${username} and ${password}
 * actionType: custom_login
 * context: web
 * needsLocator: false
 * category: Authentication
 */
export async function login(ctx: WalnutContext) {
  // ctx.args contains resolved values from ${...} placeholders in the description, in order
  const webCtx = ctx as import('./walnut').WalnutWebContext;
  await webCtx.navigate(ctx.testBaseUrl + '/login');
  await webCtx.type('[data-testid="username"]', ctx.args[0]);
  await webCtx.type('[data-testid="password"]', ctx.args[1]);
  await webCtx.click('[data-testid="submit"]');
  await webCtx.verifyTextVisible('Dashboard');
}
