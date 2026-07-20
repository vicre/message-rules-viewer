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

The application is available at `http://localhost:8080`. The frontend container
proxies `/api` requests to the backend container, so the backend is not exposed
on the host network.

The default `ALLOWED_ORIGIN` is `http://localhost:8080`. To use a different
published frontend port or hostname, set both values before starting Compose:

```sh
FRONTEND_PORT=3000 ALLOWED_ORIGIN=http://localhost:3000 docker compose up --build
```

Stop the stack with `docker compose down`.
