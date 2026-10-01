const CHECKPOINT_RECIPIENTS = {
  carga_inicio:          ["analista", "controladoria"],
  carga_fim:             ["asistente_principal", "asistente_substituta"],
  liberacao_asistente:   ["analista", "controladoria"],
  saida_portaria:        ["analista"],
  abastecimento_inicio:  ["analista"],
  retorno_fisico:        ["analista", "conferente_retorno"],
  retorno_analista_ok:   ["conferente_retorno"],
  conferencia_retorno:   ["analista", "controladoria"],
}

const DELAY_MESSAGES = {
  liberacao_asistente: (vda, motorista, minutes) => ({
    title: `ATRASO — VDA ${vda}`,
    body:  `Carro ${vda} (${motorista}) aguarda liberacao ha ${minutes} min. SLA: 10 min.`,
  }),
  abastecimento_inicio: (vda, motorista, minutes) => ({
    title: `Abastecimento longo — VDA ${vda}`,
    body:  `VDA ${vda} (${motorista}) em abastecimento ha ${minutes} min. SLA: 10 min.`,
  }),
}

export class NotificationService {
  constructor(db, redis, wsClients) {
    this.db        = db
    this.redis     = redis
    this.wsClients = wsClients
  }

  async notifyCheckpoint(vehicleId, vda, motorista, checkpoint, extraData = {}) {
    const roles = CHECKPOINT_RECIPIENTS[checkpoint] || []
    if (!roles.length) return
    const actors = await this._getActorsByRoles(roles)
    for (const actor of actors) {
      const msg = this._buildMessage(checkpoint, vda, motorista, extraData)
      await this._send(actor, vehicleId, null, msg)
    }
  }

  async notifyDelay(vehicleId, vda, motorista, checkpoint, delayMinutes, reason = null) {
    const roles  = CHECKPOINT_RECIPIENTS[checkpoint] || ["analista", "controladoria"]
    const actors = await this._getActorsByRoles(roles)
    const msgFn  = DELAY_MESSAGES[checkpoint]
    const msg = msgFn
      ? msgFn(vda, motorista, delayMinutes)
      : { title: `ATRASO — VDA ${vda}`, body: `${checkpoint} com ${delayMinutes} min de atraso.` }
    if (reason) msg.body += ` Motivo: ${reason}.`
    for (const actor of actors) {
      await this._send(actor, vehicleId, null, msg, true)
    }
  }

  async broadcastStateChange(vehicleId, vda, newState, data = {}) {
    const payload = JSON.stringify({
      type: "STATE_CHANGE", vehicleId, vda, newState,
      timestamp: new Date().toISOString(), data,
    })
    await this.redis.publish("vehicle:state", payload)
    this._broadcastWS(payload)
  }

  async startSLATimer(vehicleId, vda, motorista, checkpoint, slaMinutes) {
    const key = `sla:${vehicleId}:${checkpoint}`
    const ttl = slaMinutes * 60
    await this.redis.setEx(key, ttl, JSON.stringify({
      vehicleId, vda, motorista, checkpoint, slaMinutes,
      startedAt: new Date().toISOString(),
    }))
  }

  async cancelSLATimer(vehicleId, checkpoint) {
    await this.redis.del(`sla:${vehicleId}:${checkpoint}`)
  }

  _sendWS(actorId, payload) {
    const sockets = this.wsClients.get(actorId)
    if (!sockets) return
    const str = JSON.stringify(payload)
    for (const ws of sockets) {
      if (ws.readyState === 1) ws.send(str)
    }
  }

  _broadcastWS(payload) {
    for (const sockets of this.wsClients.values()) {
      for (const ws of sockets) {
        if (ws.readyState === 1) ws.send(payload)
      }
    }
  }

  _buildMessage(checkpoint, vda, motorista, extra) {
    const msgs = {
      carga_inicio:        { title: `Carga iniciada — VDA ${vda}`,        body: `${motorista} iniciou o carregamento.` },
      carga_fim:           { title: `Carga finalizada — VDA ${vda}`,       body: `${motorista} finalizou. Aguardando liberacao.` },
      liberacao_asistente: { title: `Carro liberado — VDA ${vda}`,         body: `${motorista} liberado para saida.` },
      saida_portaria:      { title: `Saida portaria — VDA ${vda}`,         body: `${motorista} saiu. KM: ${extra.km || "—"}.` },
      retorno_fisico:      { title: `Retorno — VDA ${vda}`,                body: `${motorista} retornou.` },
      conferencia_retorno: { title: `Conferencia concluida — VDA ${vda}`,  body: `Eferson confirma retorno fisico.` },
    }
    return msgs[checkpoint] || { title: `VDA ${vda}`, body: checkpoint }
  }

  async _send(actor, vehicleId, eventId, msg, urgent = false) {
    this._sendWS(actor.id, { type: urgent ? "ALERT" : "NOTIFICATION", ...msg, urgent, vehicleId, timestamp: new Date().toISOString() })
    if (actor.fcm_token) await this._sendFCM(actor.fcm_token, msg, urgent)
    await this.db.query(
      `INSERT INTO notifications (vehicle_id, event_id, recipient_id, recipient_name, channel, title, body, sent_at, delivered)
       VALUES ($1, $2, $3, $4, $5, $6, $7, NOW(), false)`,
      [vehicleId, eventId, actor.id, actor.name, "push_fcm", msg.title, msg.body]
    ).catch(() => {})
  }

  async _sendFCM(token, msg, urgent) {
    try {
      await fetch("https://fcm.googleapis.com/fcm/send", {
        method: "POST",
        headers: { "Authorization": `key=${process.env.FCM_SERVER_KEY}`, "Content-Type": "application/json" },
        body: JSON.stringify({ to: token, notification: { title: msg.title, body: msg.body, sound: urgent ? "alarm" : "default" }, priority: urgent ? "high" : "normal" }),
      })
    } catch {}
  }

  async _getActorsByRoles(roles) {
    const { rows } = await this.db.query(
      `SELECT id, name, role, fcm_token FROM actors WHERE role = ANY($1::actor_role_enum[]) AND is_active = TRUE`,
      [roles]
    )
    return rows
  }
}
