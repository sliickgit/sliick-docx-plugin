import { describe, expect, it } from "vitest";
import {
  base64UrlEncode,
  buildAuthorizeUrl,
  computeCodeChallenge,
  generateCodeVerifier,
} from "./pkce";

describe("generateCodeVerifier", () => {
  it("produces the requested length from the unreserved charset", () => {
    const v = generateCodeVerifier(128);
    expect(v).toHaveLength(128);
    expect(v).toMatch(/^[A-Za-z0-9\-._~]+$/);
  });

  it("rejects out-of-spec lengths", () => {
    expect(() => generateCodeVerifier(42)).toThrow();
    expect(() => generateCodeVerifier(129)).toThrow();
  });

  it("is random across calls", () => {
    expect(generateCodeVerifier()).not.toEqual(generateCodeVerifier());
  });
});

describe("computeCodeChallenge", () => {
  it("matches the RFC 7636 appendix B test vector", async () => {
    const challenge = await computeCodeChallenge(
      "dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk",
    );
    expect(challenge).toBe("E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM");
  });
});

describe("base64UrlEncode", () => {
  it("emits url-safe output without padding", () => {
    const bytes = new Uint8Array([251, 255, 254, 0, 1]).buffer;
    const out = base64UrlEncode(bytes);
    expect(out).not.toMatch(/[+/=]/);
  });
});

describe("buildAuthorizeUrl", () => {
  it("includes all required PKCE params", () => {
    const url = new URL(
      buildAuthorizeUrl({
        orgUrl: "https://acme.my.salesforce.com",
        clientId: "3MVG9xyz",
        redirectUri: "https://localhost:3000/auth-callback.html",
        codeChallenge: "abc123",
        state: "st4te",
      }),
    );
    expect(url.origin).toBe("https://acme.my.salesforce.com");
    expect(url.pathname).toBe("/services/oauth2/authorize");
    expect(url.searchParams.get("response_type")).toBe("code");
    expect(url.searchParams.get("client_id")).toBe("3MVG9xyz");
    expect(url.searchParams.get("code_challenge")).toBe("abc123");
    expect(url.searchParams.get("code_challenge_method")).toBe("S256");
    expect(url.searchParams.get("state")).toBe("st4te");
    expect(url.searchParams.get("scope")).toBe("api refresh_token openid");
  });
});
