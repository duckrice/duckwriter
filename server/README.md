# DuckWriter relay server

Lets people write a DuckWriter document together over the internet. It only
passes messages between the person sharing and those who join; the messages
are encrypted in DuckWriter with a key the server never gets, so it can't read
the documents. It stores nothing.

## Run it

Anywhere with Node.js 20 or newer:

```bash
npm install
npm start
```

It listens on port 8787 (or `PORT`). Put it behind HTTPS (most hosts do this
for you) and give DuckWriter the address as `wss://your-server.example.com`.

With Docker:

```bash
docker build -t duckwriter-relay .
docker run -p 8787:8787 duckwriter-relay
```

Hosts that work well: Fly.io, Render, Railway, any small VPS, or a Raspberry Pi
at home (with a port forwarded, or a tunnel such as Cloudflare Tunnel).

Opening `https://your-server.example.com/` in a browser shows that it is
running and how many rooms are open.
