export default async function dashboardRoutes(fastify) {

  fastify.get("/dashboard/:date", async (req, reply) => {
    const { date } = req.params
    const { rows: vehicles } = await fastify.db.query(
      `SELECT v.id, v.sequence, v.rota, v.vda, v.vehicle_type, v.motorista_name, v.equipe, v.state,
              v.paletes_realizados, v.estrados_realizados, v.km_saida, v.km_retorno,
              v.carga_inicio_ts, v.carga_fim_ts, v.liberacao_asistente_ts, v.saida_portaria_ts,
              v.retorno_fisico_ts, v.finalizado_ts, v.duracao_carga_min, v.duracao_liberacao_min,
              v.liberacao_atrasada, (v.km_retorno - v.km_saida) AS km_percorrido
       FROM vehicles v
       JOIN daily_plans dp ON dp.id = v.daily_plan_id
       WHERE dp.operation_date = $1
       ORDER BY v.sequence`,
      [date]
    )
    const byState = vehicles.reduce((acc, v) => {
      acc[v.state] = (acc[v.state] || 0) + 1
      return acc
    }, {})
    const alerts = []
    for (const v of vehicles) {
      if (v.state === "liberado_producao" && v.carga_fim_ts) {
        const diffMin = Math.floor((Date.now() - new Date(v.carga_fim_ts).getTime()) / 60000)
        if (diffMin > 10) {
          alerts.push({
            type: "LIBERACAO_ATRASADA", urgency: "high",
            vda: v.vda, rota: v.rota, motorista: v.motorista_name,
            message: `VDA ${v.vda} aguarda liberacao ha ${diffMin} min (SLA: 10 min)`,
            since: v.carga_fim_ts, diffMin,
          })
        }
      }
      if (v.state === "abastecendo") {
        alerts.push({
          type: "ABASTECIMENTO_LONGO", urgency: "medium",
          vda: v.vda, rota: v.rota, motorista: v.motorista_name,
          message: `VDA ${v.vda} em abastecimento`,
        })
      }
    }
    return { date, totalVehicles: vehicles.length, byState, alerts, vehicles }
  })

  fastify.get("/dashboard/:date/delays", async (req, reply) => {
    const { rows } = await fastify.db.query(
      `SELECT ve.id, v.vda, v.rota, v.motorista_name, ve.checkpoint, ve.actor_name,
              ve.is_delayed, ve.delay_minutes, ve.delay_reason, ve.delay_justification, ve.occurred_at
       FROM vehicle_events ve
       JOIN vehicles v ON v.id = ve.vehicle_id
       JOIN daily_plans dp ON dp.id = v.daily_plan_id
       WHERE dp.operation_date = $1 AND ve.is_delayed = TRUE
       ORDER BY ve.occurred_at DESC`,
      [req.params.date]
    )
    return rows
  })

  fastify.get("/dashboard/planned-vs-realized/:date", async (req, reply) => {
    const { rows } = await fastify.db.query(
      `SELECT * FROM v_planned_vs_realized WHERE operation_date = $1 ORDER BY sequence`,
      [req.params.date]
    )
    return rows
  })

  fastify.get("/dashboard/delays/summary/:date", async (req, reply) => {
    const { rows } = await fastify.db.query(
      `SELECT * FROM v_delay_summary WHERE operation_date = $1`,
      [req.params.date]
    )
    return rows
  })
}
