export default async function websocketRoutes(fastify) {

  fastify.get("/ws", { websocket: true }, (socket, req) => {
    const actorId = req.query.actorId || "anonymous"

    if (!fastify.wsClients.has(actorId)) {
      fastify.wsClients.set(actorId, new Set())
    }
    fastify.wsClients.get(actorId).add(socket)

    fastify.log.info({ actorId }, "WebSocket conectado")

    sendDailySnapshot(fastify, socket)

    socket.on("message", (raw) => {
      try {
        const msg = JSON.parse(raw.toString())
        if (msg.type === "PING") {
          socket.send(JSON.stringify({ type: "PONG", ts: Date.now() }))
        }
      } catch {}
    })

    socket.on("close", () => {
      const sockets = fastify.wsClients.get(actorId)
      if (sockets) {
        sockets.delete(socket)
        if (sockets.size === 0) fastify.wsClients.delete(actorId)
      }
      fastify.log.info({ actorId }, "WebSocket desconectado")
    })

    socket.on("error", (err) => {
      fastify.log.error({ err, actorId }, "WebSocket error")
    })
  })
}

async function sendDailySnapshot(fastify, socket) {
  try {
    const today = new Date().toISOString().slice(0, 10)
    const { rows } = await fastify.db.query(
      `SELECT v.id, v.sequence, v.rota, v.vda, v.vehicle_type, v.motorista_name, v.state,
              v.carga_inicio_ts, v.carga_fim_ts, v.liberacao_asistente_ts, v.saida_portaria_ts,
              v.retorno_fisico_ts, v.liberacao_atrasada, v.duracao_carga_min, v.duracao_liberacao_min
       FROM vehicles v
       JOIN daily_plans dp ON dp.id = v.daily_plan_id
       WHERE dp.operation_date = $1
       ORDER BY v.sequence`,
      [today]
    )
    if (socket.readyState === 1) {
      socket.send(JSON.stringify({
        type: "DAILY_SNAPSHOT", date: today,
        vehicles: rows, timestamp: new Date().toISOString(),
      }))
    }
  } catch (err) {
    fastify.log.error({ err }, "Erro ao enviar snapshot")
  }
}
