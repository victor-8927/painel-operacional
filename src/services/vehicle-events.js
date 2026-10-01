const STATE_TRANSITIONS = {
  aguardando_carga:         ["em_carga"],
  em_carga:                 ["liberado_producao"],
  liberado_producao:        ["em_liberacao_asistente"],
  em_liberacao_asistente:   ["liberado_portaria"],
  liberado_portaria:        ["em_rota"],
  em_rota:                  ["abastecendo", "retornado"],
  abastecendo:              ["em_rota"],
  retornado:                ["retorno_analista"],
  retorno_analista:         ["retorno_conferencia"],
  retorno_conferencia:      ["finalizado"],
  finalizado:               [],
}

export class VehicleEventService {
  constructor(db, notifications) {
    this.db            = db
    this.notifications = notifications
  }

  async recordEvent({ vehicleId, checkpoint, actorId, actorName, newState, metadata = {}, delayMinutes = null, delayReason = null, delayJustification = null }) {
    const client = await this.db.pool.connect()
    try {
      await client.query("BEGIN")
      const { rows: [vehicle] } = await client.query(
        `SELECT state, vda, motorista_name FROM vehicles WHERE id = $1 FOR UPDATE`,
        [vehicleId]
      )
      if (!vehicle) throw new Error(`Veiculo nao encontrado: ${vehicleId}`)
      const stateBefore = vehicle.state
      const allowed = STATE_TRANSITIONS[stateBefore] || []
      if (!allowed.includes(newState)) {
        throw new Error(`Transicao invalida: ${stateBefore} -> ${newState} (VDA ${vehicle.vda})`)
      }
      const isDelayed = delayMinutes > 0
      const { rows: [event] } = await client.query(
        `INSERT INTO vehicle_events
           (vehicle_id, checkpoint, actor_id, actor_name, state_before, state_after, occurred_at, metadata, is_delayed, delay_minutes, delay_reason, delay_justification)
         VALUES ($1, $2, $3, $4, $5, $6, NOW(), $7, $8, $9, $10, $11)
         RETURNING id, occurred_at`,
        [vehicleId, checkpoint, actorId, actorName, stateBefore, newState, JSON.stringify(metadata), isDelayed, delayMinutes, delayReason, delayJustification]
      )
      const tsField = this._getTimestampField(checkpoint)
      let extraUpdates = ""
      const extraParams = [newState, vehicleId]
      let paramIdx = 3
      if (metadata.km_saida)   { extraUpdates += `, km_saida = $${paramIdx++}`;           extraParams.push(metadata.km_saida) }
      if (metadata.km_retorno) { extraUpdates += `, km_retorno = $${paramIdx++}`;         extraParams.push(metadata.km_retorno) }
      if (metadata.paletes)    { extraUpdates += `, paletes_realizados = $${paramIdx++}`; extraParams.push(metadata.paletes) }
      if (metadata.estrados)   { extraUpdates += `, estrados_realizados = $${paramIdx++}`; extraParams.push(metadata.estrados) }
      if (metadata.lotes)      { extraUpdates += `, lotes = $${paramIdx++}`;              extraParams.push(metadata.lotes) }
      if (metadata.abasteceu !== undefined) { extraUpdates += `, abasteceu = $${paramIdx++}`; extraParams.push(metadata.abasteceu) }
      const tsUpdate = tsField ? `, ${tsField} = NOW()` : ""
      await client.query(
        `UPDATE vehicles SET state = $1, updated_at = NOW()${tsUpdate}${extraUpdates} WHERE id = $2`,
        extraParams
      )
      await client.query("COMMIT")
      await this.notifications.broadcastStateChange(vehicleId, vehicle.vda, newState, { checkpoint, ...metadata })
      if (isDelayed) {
        await this.notifications.notifyDelay(vehicleId, vehicle.vda, vehicle.motorista_name, checkpoint, delayMinutes, delayReason)
      } else {
        await this.notifications.notifyCheckpoint(vehicleId, vehicle.vda, vehicle.motorista_name, checkpoint, metadata)
      }
      return { eventId: event.id, occurredAt: event.occurred_at, stateBefore, stateAfter: newState, isDelayed }
    } catch (err) {
      await client.query("ROLLBACK")
      throw err
    } finally {
      client.release()
    }
  }

  _getTimestampField(checkpoint) {
    const map = {
      carga_inicio:        "carga_inicio_ts",
      carga_fim:           "carga_fim_ts",
      liberacao_asistente: "liberacao_asistente_ts",
      saida_portaria:      "saida_portaria_ts",
      retorno_fisico:      "retorno_fisico_ts",
    }
    return map[checkpoint] || null
  }

  async getHistory(vehicleId) {
    const { rows } = await this.db.query(
      `SELECT id, checkpoint, actor_name, state_before, state_after, occurred_at, metadata, is_delayed, delay_minutes, delay_reason, delay_justification
       FROM vehicle_events WHERE vehicle_id = $1 ORDER BY occurred_at ASC`,
      [vehicleId]
    )
    return rows
  }
}
