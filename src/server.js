import "dotenv/config"
import Fastify from "fastify"
import fastifyWebsocket from "@fastify/websocket"
import fastifyCors from "@fastify/cors"

import dbPlugin    from "./plugins/db.js"
import redisPlugin from "./plugins/redis.js"

import { NotificationService } from "./services/notifications.js"
import { VehicleEventService } from "./services/vehicle-events.js"

import plansRoutes     from "./routes/plans.js"
import vehiclesRoutes  from "./routes/vehicles.js"
import websocketRoutes from "./routes/websocket.js"
import dashboardRoutes from "./routes/dashboard.js"

const fastify = Fastify({
  logger: {
    level: process.env.LOG_LEVEL || "info",
    transport: process.env.NODE_ENV !== "production"
      ? { target: "pino-pretty", options: { colorize: true } }
      : undefined,
  },
})

await fastify.register(fastifyCors, { origin: process.env.CORS_ORIGIN || "*" })
await fastify.register(fastifyWebsocket)
await fastify.register(dbPlugin)
await fastify.register(redisPlugin)

fastify.decorate("wsClients", new Map())

const notifications = new NotificationService(fastify.db, fastify.redis, fastify.wsClients)
const vehicleEvents = new VehicleEventService(fastify.db, notifications)

fastify.decorate("notifications", notifications)
fastify.decorate("vehicleEvents", vehicleEvents)

const subscriber = fastify.redis.duplicate()
await subscriber.connect()
await subscriber.subscribe("vehicle:state", (message) => {
  for (const sockets of fastify.wsClients.values()) {
    for (const ws of sockets) {
      if (ws.readyState === 1) ws.send(message)
    }
  }
})

await fastify.register(websocketRoutes)
await fastify.register(plansRoutes,    { prefix: "/api" })
await fastify.register(vehiclesRoutes, { prefix: "/api" })
await fastify.register(dashboardRoutes,{ prefix: "/api" })

fastify.get("/health", async () => ({ status: "ok", ts: new Date().toISOString() }))

const PORT = parseInt(process.env.PORT || "3001")
const HOST = process.env.HOST || "0.0.0.0"

try {
  await fastify.listen({ port: PORT, host: HOST })
  fastify.log.info(`Painel Operacional rodando em http://${HOST}:${PORT}`)
} catch (err) {
  fastify.log.error(err)
  process.exit(1)
}
