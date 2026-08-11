# Xinghuo Office

A self-hosted, browser-local Office editor for DOCX, XLSX, and PPTX files.

Documents are passed to the editor in browser memory and processed locally with ONLYOFFICE Web Apps and x2t WebAssembly. The production image does not run DocumentServer, upload documents to an Office backend, or require the former `office-editor.ziziyi.com` service.

## Security model

- Editor scripts, fonts, workers, and WebAssembly are served from the same origin.
- Third-party Office plugins and cloud-drive integrations are disabled.
- A restrictive Content Security Policy blocks third-party network requests.
- Builds fail if known former-operator or analytics domains remain in generated assets.
- The Xinghuo integration uses a strict-origin `postMessage` bridge and transfers document data as `ArrayBuffer`.

## Build

Requirements: Docker with access to the official images below.

```text
onlyoffice/documentserver:9.4.0.1
node:22-alpine
caddy:2-alpine
```

Build the production image:

```bash
docker build \
  --build-arg DS_VERSION=9.4.0.1 \
  --build-arg HASH=1 \
  -t xinghuo-office:9.4.0.1-2 \
  .
```

The ONLYOFFICE image is used only as a build-stage source for static resources. It is not running in the final container.

Run locally:

```bash
docker run -d \
  --name xinghuo-office \
  --restart unless-stopped \
  -p 127.0.0.1:18080:80 \
  xinghuo-office:9.4.0.1-2
```

## 1Panel Compose

Clone this repository to `/opt/office-website`, build the image with the command above, then use:

```yaml
services:
  office:
    image: xinghuo-office:9.4.0.1-2
    container_name: xinghuo-office
    restart: unless-stopped
    ports:
      - "127.0.0.1:18080:80"
```

Configure a 1Panel website reverse proxy to `http://127.0.0.1:18080` and enable HTTPS. A domain with HTTPS is strongly recommended; an HTTP IP endpoint will be blocked when embedded by an HTTPS Xinghuo web client.

## Development

```bash
pnpm install
pnpm dev
```

The source and ONLYOFFICE Community components are distributed under AGPL-3.0-compatible terms. Preserve the corresponding notices and publish source changes when providing the modified application over a network.
