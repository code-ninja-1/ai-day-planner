import { ConfidentialClientApplication } from "@azure/msal-node";
import type { Request } from "express";
import { createRemoteJWKSet, jwtVerify } from "jose";
import { env, microsoftAuthority, microsoftIssuer } from "../env.js";

const jwks = createRemoteJWKSet(new URL(`${microsoftAuthority}/discovery/v2.0/keys`));

const confidentialClient = new ConfidentialClientApplication({
  auth: {
    clientId: env.microsoftClientId,
    clientSecret: env.microsoftClientSecret,
    authority: microsoftAuthority
  }
});

const graphScopes = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Calendars.Read"
];

export interface MicrosoftSession {
  accessToken: string;
  accountLabel: string | null;
  displayName: string | null;
  oid: string | null;
}

function getBearerToken(request: Request) {
  const header = request.headers.authorization;
  if (!header?.startsWith("Bearer ")) {
    return null;
  }
  return header.slice("Bearer ".length).trim();
}

export async function getOptionalMicrosoftSession(request: Request) {
  const accessToken = getBearerToken(request);
  if (!accessToken) {
    return null;
  }

  const { payload } = await jwtVerify(accessToken, jwks, {
    issuer: microsoftIssuer,
    audience: [env.microsoftApiAudience, env.microsoftClientId].filter(Boolean)
  });

  return {
    accessToken,
    accountLabel:
      (payload.preferred_username as string | undefined) ??
      (payload.upn as string | undefined) ??
      null,
    displayName: (payload.name as string | undefined) ?? null,
    oid: (payload.oid as string | undefined) ?? null
  } satisfies MicrosoftSession;
}

export async function acquireGraphTokenOnBehalfOf(session: MicrosoftSession) {
  if (!env.microsoftClientId || !env.microsoftClientSecret) {
    throw new Error("Microsoft OBO is not configured in apps/api/.env");
  }

  const result = await confidentialClient.acquireTokenOnBehalfOf({
    oboAssertion: session.accessToken,
    scopes: graphScopes
  });

  if (!result?.accessToken) {
    throw new Error("Failed to acquire Microsoft Graph token on behalf of the user");
  }

  return result.accessToken;
}
