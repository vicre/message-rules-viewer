# Message Rules Viewer

## Run with Docker Compose

1. Create the backend environment file and add the Microsoft Entra credentials:

   ```sh
   cp server/.env.example server/.env
   ```

2. Start the frontend and backend:

   ```sh
   docker compose up --build
   ```

The application is available at `http://localhost:8080`. The `proxy` container
is the only service published on the host. It serves the frontend at `/` and
forwards `/api/` requests to the backend, so the frontend and backend are not
exposed on the host network.

The default `ALLOWED_ORIGIN` is `http://localhost:8080`. To use a different
published frontend port or hostname, set both values before starting Compose:

```sh
FRONTEND_PORT=3000 ALLOWED_ORIGIN=http://localhost:3000 docker compose up --build
```

Stop the stack with `docker compose down`.

## Deploy behind Traefik

Publish only the `proxy` service to Traefik. Traefik should terminate TLS for
`mailbox-settings-premises.security.ait.dtu.dk` and forward HTTP traffic to the
proxy container on port `80`. The proxy then routes browser requests as follows:

| Request path | Upstream |
| --- | --- |
| `/` | `frontend:80` |
| `/api/` | `backend:3001` |

Set the backend's allowed origin to the public HTTPS URL when deploying:

```sh
ALLOWED_ORIGIN=https://mailbox-settings-premises.security.ait.dtu.dk docker compose up --build -d
```

Configure the Traefik service target as `proxy:80` (or the proxy's published
host port when Traefik runs outside this Compose network). Do not publish the
`frontend` or `backend` service directly.
