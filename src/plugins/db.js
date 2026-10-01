import fp from "fastify-plugin"
import pg from "pg"

const { Pool } = pg

async function dbPlugin(fastify, opts) {
  const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    max: 10,
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 5000,
  })

  pool.on("error", (err) => {
    fastify.log.error({ err }, "PostgreSQL pool error")
  })

  fastify.decorate("db", {
    query: (text, params) => pool.query(text, params),
    pool,
  })

  fastify.addHook("onClose", async () => {
    await pool.end()
    fastify.log.info("PostgreSQL pool closed")
  })
}

export default fp(dbPlugin, { name: "db" })
