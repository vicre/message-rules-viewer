import "dotenv/config";
import express from "express";
import cors from "cors";
import {
  createRemoteJWKSet,
  jwtVerify,
} from "jose";

const app = express();

const PORT = Number(process.env.PORT ?? 3001);

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const API_AUDIENCE =
  process.env.API_AUDIENCE ??
  `api://${CLIENT_ID}`;

const REQUIRED_SCOPE =
  process.env.REQUIRED_SCOPE ??
  "access_as_user";

const ALLOWED_ORIGIN =
  process.env.ALLOWED_ORIGIN ??
  "http://localhost:5173";

const requiredVariables = [
  "TENANT_ID",
  "CLIENT_ID",
  "CLIENT_SECRET",
];

for (const variable of requiredVariables) {
  if (!process.env[variable]) {
    throw new Error(
      `Missing required environment variable: ${variable}`,
    );
  }
}

app.use(
  cors({
    origin: ALLOWED_ORIGIN,
    methods: ["GET"],
    allowedHeaders: [
      "Authorization",
      "Content-Type",
    ],
  }),
);

app.use(express.json());

const issuer =
  `https://login.microsoftonline.com/${TENANT_ID}/v2.0`;

const jwks = createRemoteJWKSet(
  new URL(
    `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`,
  ),
);

/**
 * Verifies the Entra access token sent by the Vite frontend.
 */
async function requireAuthentication(
  request,
  response,
  next,
) {
  const authorization =
    request.headers.authorization;

  if (
    !authorization ||
    !authorization.startsWith("Bearer ")
  ) {
    return response.status(401).json({
      error: "Missing bearer access token.",
    });
  }

  const token =
    authorization.slice("Bearer ".length).trim();

  if (!token) {
    return response.status(401).json({
      error: "Missing bearer access token.",
    });
  }

  try {
    const { payload } = await jwtVerify(
      token,
      jwks,
      {
        issuer,
        audience: API_AUDIENCE,
        algorithms: ["RS256"],
      },
    );

    const scopes =
      typeof payload.scp === "string"
        ? payload.scp.split(" ")
        : [];

    if (!scopes.includes(REQUIRED_SCOPE)) {
      return response.status(403).json({
        error:
          `The token does not contain the required ` +
          `scope: ${REQUIRED_SCOPE}.`,
      });
    }

    if (payload.tid !== TENANT_ID) {
      return response.status(403).json({
        error: "The token belongs to another tenant.",
      });
    }

    request.auth = {
      objectId: payload.oid,
      tenantId: payload.tid,
      username:
        payload.preferred_username ??
        payload.upn ??
        null,
      name: payload.name ?? null,
      scopes,
      claims: payload,
    };

    next();
  } catch (error) {
    console.error(
      "Access-token validation failed:",
      error,
    );

    return response.status(401).json({
      error: "Invalid or expired access token.",
    });
  }
}

async function acquireGraphToken() {
  const tokenEndpoint =
    `https://login.microsoftonline.com/` +
    `${encodeURIComponent(TENANT_ID)}` +
    `/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const response = await fetch(tokenEndpoint, {
    method: "POST",
    headers: {
      "Content-Type":
        "application/x-www-form-urlencoded",
    },
    body,
  });

  const result =
    await response.json().catch(() => ({}));

  if (!response.ok) {
    const error = new Error(
      result.error_description ||
        result.error ||
        "Token request failed",
    );

    error.status = response.status;
    throw error;
  }

  if (!result.access_token) {
    throw new Error(
      "Microsoft Entra did not return an access token.",
    );
  }

  return result.access_token;
}

function validateUserPrincipalName(value) {
  if (typeof value !== "string") {
    return null;
  }

  const upn = value.trim().toLowerCase();

  if (!/^[a-z0-9._'-]+@dtu\.dk$/i.test(upn)) {
    return null;
  }

  return upn;
}

/**
 * Optional endpoint for testing authentication.
 */
app.get(
  "/api/me",
  requireAuthentication,
  (request, response) => {
    return response.json({
      authenticated: true,
      user: {
        username: request.auth.username,
        name: request.auth.name,
        objectId: request.auth.objectId,
      },
    });
  },
);

app.get(
  "/api/message-rules",
  requireAuthentication,
  async (request, response) => {
    const upn =
      validateUserPrincipalName(
        request.query.upn,
      );

    if (!upn) {
      return response.status(400).json({
        error:
          "Enter a valid DTU user principal name, " +
          "for example user@dtu.dk.",
      });
    }

    try {
      console.log(
        `${request.auth.username ?? "Unknown user"} ` +
          `queried mailbox rules for ${upn}`,
      );

      const graphAccessToken =
        await acquireGraphToken();

      const graphEndpoint =
        `https://graph.microsoft.com/v1.0/users/` +
        `${encodeURIComponent(upn)}` +
        `/mailFolders/inbox/messageRules`;

      const graphResponse = await fetch(
        graphEndpoint,
        {
          headers: {
            Authorization:
              `Bearer ${graphAccessToken}`,
            Accept: "application/json",
          },
        },
      );

      const result =
        await graphResponse
          .json()
          .catch(() => ({}));

      if (!graphResponse.ok) {
        return response
          .status(graphResponse.status)
          .json({
            error:
              result.error?.message ||
              `Microsoft Graph returned HTTP ` +
                `${graphResponse.status}.`,
            graphError: result.error,
          });
      }

      return response.json({
        requestedBy:
          request.auth.username,
        userPrincipalName: upn,
        rules: result.value ?? [],
      });
    } catch (error) {
      console.error(error);

      return response
        .status(error.status || 500)
        .json({
          error:
            error.message ||
            "Unexpected server error.",
        });
    }
  },
);

app.listen(PORT, () => {
  console.log(
    `API listening on http://localhost:${PORT}`,
  );

  console.log(
    `Expected audience: ${API_AUDIENCE}`,
  );

  console.log(
    `Required scope: ${REQUIRED_SCOPE}`,
  );
});