import fp from "fastify-plugin"
import { createClient } from "redis"

async function redisPlugin(fastify, opts) {
  const client = createClient({
    url: process.env.REDIS_URL,
    socket: {
      tls: process.env.REDIS_URL?.startsWith("rediss"),
      rejectUnauthorized: false
    }
  })

  client.on("error", (err) => {
    fastify.log.error({ err }, "Redis client error")
  })

  await client.connect()
  fastify.log.info("Redis connected")

  fastify.decorate("redis", client)

  fastify.addHook("onClose", async () => {
    await client.quit()
    fastify.log.info("Redis connection closed")
  })
}

export default fp(redisPlugin, { name: "redis" })
