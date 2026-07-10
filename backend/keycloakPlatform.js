'use strict';
// =====================================================
// Единый вход через платформу (Keycloak) — проверка access-токена по JWKS.
// Реализация на jwks-rsa + jsonwebtoken (CommonJS, без ESM). Проверяем подпись,
// iss, exp и azp (public-клиент: aud обычно "account", поэтому aud НЕ требуем).
// Токен НИГДЕ не логируется и не сохраняется — только причины отказа.
// =====================================================
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const PLATFORM_SSO = ['1', 'true', 'yes'].includes(String(process.env.PLATFORM_SSO || 'false').toLowerCase());
const KEYCLOAK_URL = (process.env.KEYCLOAK_URL || 'https://keycloak-ashinoff.amvera.io').replace(/\/+$/, '');
const KEYCLOAK_REALM = process.env.KEYCLOAK_REALM || 'platform';
const KEYCLOAK_AZP = process.env.KEYCLOAK_AZP || 'web-desktop';
// Одна realm-роль доступа к приложению (паттерн платформы: <app>-user).
const ACCESS_ROLE = process.env.ACCESS_ROLE || 'resm-user';
const PLATFORM_ORIGIN = process.env.PLATFORM_ORIGIN || 'https://sue-system-ashinoff.amvera.io';

const ISSUER = `${KEYCLOAK_URL}/realms/${KEYCLOAK_REALM}`;
const JWKS_URI = `${ISSUER}/protocol/openid-connect/certs`;

// JWKS кэшируется в клиенте (сам перечитывает ключи), Keycloak не дёргается на
// каждый запрос.
const client = jwksClient({
  jwksUri: JWKS_URI,
  cache: true,
  cacheMaxEntries: 5,
  cacheMaxAge: 3600000, // 1 час
  rateLimit: true,
  jwksRequestsPerMinute: 10,
});

function getKey(header, callback) {
  client.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

// Проверить Keycloak access-токен. Резолвит claims или бросает Error с
// безопасным сообщением (без самого токена).
function verifyToken(token) {
  return new Promise((resolve, reject) => {
    jwt.verify(token, getKey, { algorithms: ['RS256'], issuer: ISSUER }, (err, claims) => {
      if (err) return reject(new Error(`invalid token (${err.name})`));
      if (!claims || claims.azp !== KEYCLOAK_AZP) {
        return reject(new Error(`unexpected azp: ${claims && claims.azp}`));
      }
      resolve(claims);
    });
  });
}

function identityFromClaims(claims) {
  return {
    keycloakId: claims.sub,
    email: claims.email,
    name: claims.name,
    username: claims.preferred_username,
    roles: (claims.realm_access && claims.realm_access.roles) || [],
  };
}

// Есть ли у пользователя роль доступа к приложению.
function hasAccess(roles) {
  return Array.isArray(roles) && roles.includes(ACCESS_ROLE);
}

module.exports = {
  PLATFORM_SSO, KEYCLOAK_URL, KEYCLOAK_REALM, KEYCLOAK_AZP, ACCESS_ROLE,
  PLATFORM_ORIGIN, ISSUER, JWKS_URI,
  verifyToken, identityFromClaims, hasAccess,
};
