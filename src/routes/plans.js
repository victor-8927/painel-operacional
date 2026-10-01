export default async function plansRoutes(fastify) {

  fastify.post("/plans/import", async (req, reply) => {
    const plan = req.body
    if (!plan?.vehicles?.length) {
      return reply.code(400).send({ error: "JSON invalido — sem veiculos" })
    }
    const client = await fastify.db.pool.connect()
    try {
      await client.query("BEGIN")
      const dateMatch = plan.source_date?.match(/(\d{2})\/(\d{2})\/(\d{4})/)
      if (!dateMatch) return reply.code(400).send({ error: "Data nao encontrada no JSON" })
      const operationDate = `${dateMatch[3]}-${dateMatch[2]}-${dateMatch[1]}`
      const { rows: existing } = await client.query(
        `SELECT id FROM daily_plans WHERE operation_date = $1`, [operationDate]
      )
      if (existing.length) {
        await client.query("ROLLBACK")
        return reply.code(409).send({ error: `Plano para ${operationDate} ja foi importado`, planId: existing[0].id })
      }
      const { rows: [dp] } = await client.query(
        `INSERT INTO daily_plans (operation_date, sheet_name, source_file, raw_json, total_vehicles)
         VALUES ($1, $2, $3, $4, $5) RETURNING id`,
        [operationDate, plan.sheet_name, plan.source_date, JSON.stringify(plan), plan.total_vehicles]
      )
      const planId = dp.id
      const vehicleIds = []
      for (const v of plan.vehicles) {
        const { rows: [veh] } = await client.query(
          `INSERT INTO vehicles (daily_plan_id, sequence, rota, vda, vehicle_type, capacity_kg, motorista_name, equipe, state)
           VALUES ($1, $2, $3, $4, $5::vehicle_type_enum, $6, $7, $8, 'aguardando_carga') RETURNING id`,
          [planId, v.sequence, v.rota, v.vda, v.vehicle_type === "3/4" ? "3/4" : v.vehicle_type, v.capacity_kg, v.motorista, v.equipe]
        )
        vehicleIds.push({ vda: v.vda, id: veh.id })
        for (const vol of v.volumes) {
          await client.query(
            `INSERT INTO vehicle_volumes
               (vehicle_id, volume_type, top_sankhya, oc_exclusive, return_rule, oc_number, oc_pending,
                planned_kg3, planned_kg5, planned_kg10, planned_kg20, planned_kg40, planned_kg50, planned_peso_kg, obs)
             VALUES ($1, $2::volume_type_enum, $3, $4, $5::return_rule_enum, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15)`,
            [veh.id, vol.type, vol.top_sankhya, vol.oc_exclusive, vol.return_rule, vol.oc_number, vol.oc_pending,
             vol.skus.kg3, vol.skus.kg5, vol.skus.kg10, vol.skus.kg20, vol.skus.kg40, vol.skus.kg50, vol.peso_kg, vol.obs]
          )
        }
      }
      await client.query("COMMIT")
      return reply.code(201).send({ planId, operationDate, totalVehicles: plan.total_vehicles, vehicles: vehicleIds })
    } catch (err) {
      await client.query("ROLLBACK")
      fastify.log.error({ err }, "Erro ao importar plano")
      return reply.code(500).send({ error: err.message })
    } finally {
      client.release()
    }
  })

  fastify.get("/plans/:date", async (req, reply) => {
    const { rows: [plan] } = await fastify.db.query(
      `SELECT id, operation_date, sheet_name, total_vehicles, imported_at FROM daily_plans WHERE operation_date = $1`,
      [req.params.date]
    )
    if (!plan) return reply.code(404).send({ error: "Plano nao encontrado" })
    return plan
  })

  fastify.get("/plans/:date/vehicles", async (req, reply) => {
    const { rows } = await fastify.db.query(
      `SELECT v.id, v.sequence, v.rota, v.vda, v.vehicle_type, v.capacity_kg, v.motorista_name, v.equipe, v.state,
              v.paletes_realizados, v.estrados_realizados, v.lotes, v.km_saida, v.km_retorno,
              v.carga_inicio_ts, v.carga_fim_ts, v.liberacao_asistente_ts, v.saida_portaria_ts,
              v.retorno_fisico_ts, v.finalizado_ts, v.duracao_carga_min, v.duracao_liberacao_min, v.liberacao_atrasada
       FROM vehicles v
       JOIN daily_plans dp ON dp.id = v.daily_plan_id
       WHERE dp.operation_date = $1
       ORDER BY v.sequence`,
      [req.params.date]
    )
    return rows
  })
}
