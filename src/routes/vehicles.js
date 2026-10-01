export default async function vehiclesRoutes(fastify) {

  async function getVehicle(id) {
    const { rows: [v] } = await fastify.db.query(
      `SELECT id, vda, rota, state, motorista_name, carga_fim_ts FROM vehicles WHERE id = $1`, [id]
    )
    return v
  }

  fastify.post("/vehicles/:id/carga-inicio", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (vehicle.state !== "aguardando_carga") return reply.code(409).send({ error: `Estado atual: ${vehicle.state}` })
    const result = await fastify.vehicleEvents.recordEvent({
      vehicleId: vehicle.id, checkpoint: "carga_inicio",
      actorId: req.body.actorId, actorName: req.body.actorName || "Edinaldo Palmas",
      newState: "em_carga", metadata: {},
    })
    await fastify.redis.set(`carga:${vehicle.id}:inicio`, new Date().toISOString())
    return { ...result, vda: vehicle.vda }
  })

  fastify.post("/vehicles/:id/carga-fim", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (vehicle.state !== "em_carga") return reply.code(409).send({ error: `Estado atual: ${vehicle.state}` })
    const { paletes, estrados, lotes, paletesConfirmadoZero, lotesConfirmadoVazio } = req.body
    if (paletes === undefined && !paletesConfirmadoZero) return reply.code(400).send({ error: "Informe a quantidade de paletes.", field: "paletes" })
    if ((!lotes || lotes.length === 0) && !lotesConfirmadoVazio) return reply.code(400).send({ error: "Informe pelo menos um lote.", field: "lotes" })
    const inicioStr = await fastify.redis.get(`carga:${vehicle.id}:inicio`)
    const duracaoMin = inicioStr ? Math.floor((Date.now() - new Date(inicioStr).getTime()) / 60000) : null
    const result = await fastify.vehicleEvents.recordEvent({
      vehicleId: vehicle.id, checkpoint: "carga_fim",
      actorId: req.body.actorId, actorName: req.body.actorName || "Edinaldo Palmas",
      newState: "liberado_producao", metadata: { paletes, estrados, lotes, duracaoMin },
    })
    await fastify.redis.setEx(`sla:${vehicle.id}:liberacao_asistente`, 600,
      JSON.stringify({ vehicleId: vehicle.id, vda: vehicle.vda, motorista: vehicle.motorista_name, startedAt: new Date().toISOString() })
    )
    return { ...result, vda: vehicle.vda, duracaoMin }
  })

  fastify.post("/vehicles/:id/pular", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (!req.body.motivo) return reply.code(400).send({ error: "Motivo obrigatorio para pular sequencia." })
    await fastify.db.query(
      `INSERT INTO vehicle_events (vehicle_id, checkpoint, actor_id, actor_name, state_before, state_after, occurred_at, metadata)
       VALUES ($1, 'carga_inicio', $2, $3, $4, $4, NOW(), $5)`,
      [vehicle.id, req.body.actorId, req.body.actorName || "Edinaldo Palmas", vehicle.state,
       JSON.stringify({ pulado: true, motivo: req.body.motivo, motivo_livre: req.body.motivo_livre })]
    )
    await fastify.notifications.broadcastStateChange(vehicle.id, vehicle.vda, "pulado", { motivo: req.body.motivo })
    return { vda: vehicle.vda, pulado: true, motivo: req.body.motivo }
  })

  fastify.post("/vehicles/:id/liberacao", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (vehicle.state !== "liberado_producao") return reply.code(409).send({ error: `Estado atual: ${vehicle.state}` })
    const slaKey  = `sla:${vehicle.id}:liberacao_asistente`
    const slaData = await fastify.redis.get(slaKey)
    let delayMinutes = 0, isDelayed = false
    if (slaData) {
      const { startedAt } = JSON.parse(slaData)
      const diffMin = Math.floor((Date.now() - new Date(startedAt).getTime()) / 60000)
      delayMinutes = Math.max(0, diffMin - 10)
      isDelayed = delayMinutes > 0
    }
    if (isDelayed && !req.body.justificativa) return reply.code(400).send({ error: "Justificativa obrigatoria — SLA de 10 minutos excedido.", delayMinutes, field: "justificativa" })
    const result = await fastify.vehicleEvents.recordEvent({
      vehicleId: vehicle.id, checkpoint: "liberacao_asistente",
      actorId: req.body.actorId, actorName: req.body.actorName,
      newState: "liberado_portaria",
      delayMinutes: isDelayed ? delayMinutes : null,
      delayJustification: req.body.justificativa,
    })
    await fastify.redis.del(slaKey)
    return { ...result, vda: vehicle.vda, isDelayed, delayMinutes }
  })

  fastify.post("/vehicles/:id/saida-portaria", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (vehicle.state !== "liberado_portaria") return reply.code(409).send({ error: `Estado atual: ${vehicle.state}` })
    if (!req.body.km_saida) return reply.code(400).send({ error: "KM de saida obrigatorio.", field: "km_saida" })
    const result = await fastify.vehicleEvents.recordEvent({
      vehicleId: vehicle.id, checkpoint: "saida_portaria",
      actorId: req.body.actorId, actorName: req.body.actorName,
      newState: "em_rota", metadata: { km_saida: req.body.km_saida },
    })
    return { ...result, vda: vehicle.vda }
  })

  fastify.post("/vehicles/:id/abastecimento", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (req.body.iniciando) {
      await fastify.vehicleEvents.recordEvent({ vehicleId: vehicle.id, checkpoint: "abastecimento_inicio", actorId: req.body.actorId, actorName: req.body.actorName, newState: "abastecendo", metadata: { abasteceu: true } })
      await fastify.redis.setEx(`sla:${vehicle.id}:abastecimento`, 600, JSON.stringify({ vehicleId: vehicle.id, vda: vehicle.vda, startedAt: new Date().toISOString() }))
    } else {
      const slaData = await fastify.redis.get(`sla:${vehicle.id}:abastecimento`)
      let delayMinutes = 0
      if (slaData) { const { startedAt } = JSON.parse(slaData); delayMinutes = Math.max(0, Math.floor((Date.now() - new Date(startedAt).getTime()) / 60000) - 10) }
      await fastify.vehicleEvents.recordEvent({ vehicleId: vehicle.id, checkpoint: "abastecimento_fim", actorId: req.body.actorId, actorName: req.body.actorName, newState: "em_rota", delayMinutes: delayMinutes || null })
      await fastify.redis.del(`sla:${vehicle.id}:abastecimento`)
    }
    return { vda: vehicle.vda, abastecendo: req.body.iniciando }
  })

  fastify.post("/vehicles/:id/retorno", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (!["em_rota","abastecendo"].includes(vehicle.state)) return reply.code(409).send({ error: `Estado atual: ${vehicle.state}` })
    if (!req.body.km_retorno) return reply.code(400).send({ error: "KM de retorno obrigatorio.", field: "km_retorno" })
    const result = await fastify.vehicleEvents.recordEvent({
      vehicleId: vehicle.id, checkpoint: "retorno_fisico",
      actorId: req.body.actorId, actorName: req.body.actorName,
      newState: "retornado", metadata: { km_retorno: req.body.km_retorno },
    })
    return { ...result, vda: vehicle.vda }
  })

  fastify.post("/vehicles/:id/retorno-analista", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (vehicle.state !== "retornado") return reply.code(409).send({ error: `Estado atual: ${vehicle.state}` })
    const client = await fastify.db.pool.connect()
    try {
      await client.query("BEGIN")
      for (const vol of (req.body.volumes || [])) {
        await client.query(
          `UPDATE vehicle_volumes SET returned_kg3=$1, returned_kg5=$2, returned_kg10=$3, returned_kg20=$4, returned_kg40=$5, returned_kg50=$6, return_confirmed_at=NOW(), return_confirmed_by=$7 WHERE id=$8 AND vehicle_id=$9`,
          [vol.returned_kg3||0, vol.returned_kg5||0, vol.returned_kg10||0, vol.returned_kg20||0, vol.returned_kg40||0, vol.returned_kg50||0, req.body.actorId, vol.volumeId, vehicle.id]
        )
      }
      await client.query("COMMIT")
    } catch (err) { await client.query("ROLLBACK"); throw err } finally { client.release() }
    const result = await fastify.vehicleEvents.recordEvent({
      vehicleId: vehicle.id, checkpoint: "retorno_analista_ok",
      actorId: req.body.actorId, actorName: req.body.actorName || "Victor Mosquera",
      newState: "retorno_analista",
    })
    return { ...result, vda: vehicle.vda }
  })

  fastify.post("/vehicles/:id/conferencia", async (req, reply) => {
    const vehicle = await getVehicle(req.params.id)
    if (!vehicle) return reply.code(404).send({ error: "Veiculo nao encontrado" })
    if (vehicle.state !== "retorno_analista") return reply.code(409).send({ error: `Estado atual: ${vehicle.state}` })
    const b = req.body
    await fastify.db.query(
      `INSERT INTO return_conferences (vehicle_id, conferente_id, paletes_recebidos, estrados_recebidos, sacos_recebidos_kg3, sacos_recebidos_kg5, sacos_recebidos_kg10, sacos_recebidos_kg20, sacos_recebidos_kg40, sacos_recebidos_kg50, has_divergence, divergence_notes, confirmed_at)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,NOW())`,
      [vehicle.id, b.actorId, b.paletes_recebidos||0, b.estrados_recebidos||0, b.sacos_recebidos_kg3||0, b.sacos_recebidos_kg5||0, b.sacos_recebidos_kg10||0, b.sacos_recebidos_kg20||0, b.sacos_recebidos_kg40||0, b.sacos_recebidos_kg50||0, b.has_divergence||false, b.divergence_notes||null]
    )
    const result = await fastify.vehicleEvents.recordEvent({
      vehicleId: vehicle.id, checkpoint: "conferencia_retorno",
      actorId: b.actorId, actorName: b.actorName || "Eferson dos Santos",
      newState: b.has_divergence ? "retorno_conferencia" : "finalizado",
      metadata: { has_divergence: b.has_divergence },
    })
    return { ...result, vda: vehicle.vda, hasDivergence: b.has_divergence }
  })

  fastify.get("/vehicles/:id/history", async (req, reply) => {
    return fastify.vehicleEvents.getHistory(req.params.id)
  })

  fastify.get("/vehicles/:id/volumes", async (req, reply) => {
    const { rows } = await fastify.db.query(
      `SELECT id, volume_type, top_sankhya, oc_number, oc_pending,
              planned_kg3, planned_kg5, planned_kg10, planned_kg20, planned_kg40, planned_kg50,
              returned_kg3, returned_kg5, returned_kg10, returned_kg20, returned_kg40, returned_kg50,
              return_rule, obs, return_confirmed_at
       FROM vehicle_volumes WHERE vehicle_id = $1 ORDER BY id`,
      [req.params.id]
    )
    return rows
  })
}
