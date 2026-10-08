import XLSX from "xlsx"

export default async function plansRoutes(fastify) {

  // ─── Helper: parseia buffer xlsx → preview de veículos ─────────────────────
  function parseXlsxBuffer(buffer, targetDate) {
    const wb = XLSX.read(buffer, { type: "buffer", raw: true })

    // Determinar data de operação
    let operationDate = targetDate
    if (!operationDate) {
      const now = new Date(Date.now() - 3 * 60 * 60 * 1000)
      operationDate = now.toISOString().split("T")[0]
    }

    // Nome da sheet DD.MM.YYYY
    const [y, m, d] = operationDate.split("-")
    const sheetName = `${d}.${m}.${y}`
    const ws = wb.Sheets[sheetName]
    if (!ws) {
      const available = wb.SheetNames.join(", ")
      throw Object.assign(new Error(`Sheet '${sheetName}' nao encontrada`), { code: 404, sheets_available: available })
    }

    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: false })

    let headerRow = -1
    for (let i = 0; i < Math.min(raw.length, 10); i++) {
      if (raw[i]?.some(c => typeof c === "string" && c.toUpperCase().includes("VDA"))) {
        headerRow = i
        break
      }
    }
    if (headerRow === -1) throw Object.assign(new Error("Cabecalho com 'VDA' nao encontrado"), { code: 422 })

    const numOf = (v) => { const n = Number(v); return isNaN(n) ? 0 : Math.round(n) }
    const toKey = (v) => {
      if (v === null || v === undefined) return null
      const s = String(v).trim().toUpperCase()
      if (s.includes("ACCELO") || s.includes("ACELO")) return "ACCELO"
      const digits = s.replace(/\D/g, "")
      return digits ? String(parseInt(digits, 10)) : null
    }

    // Agrupa por VDA: cada VDA pode ter múltiplas linhas (PRE-VENDA, TROCAS, BONIFICAÇÃO…)
    // Somamos os SKUs de todas as linhas e usamos motorista/rota/cap da primeira linha do VDA
    const vdaMap = new Map()   // key → { sequence, vda, rota, motorista_name, capacity_kg, kg3..kg50, obs }
    const vdaOrder = []        // ordem de inserção para manter a sequência original

    for (let i = headerRow + 1; i < raw.length; i++) {
      const r = raw[i]
      if (!r || r.every(c => c === null)) continue
      const vdaKey = toKey(r[1])
      if (!vdaKey) continue

      const kg3  = numOf(r[10]), kg5  = numOf(r[11]), kg10 = numOf(r[12])
      const kg20 = numOf(r[13]), kg40 = numOf(r[14]), kg50 = numOf(r[15])

      if (vdaMap.has(vdaKey)) {
        // VDA já existe → acumula SKUs (PRE-VENDA + TROCAS + outras linhas)
        const entry = vdaMap.get(vdaKey)
        entry.planned_kg3  += kg3
        entry.planned_kg5  += kg5
        entry.planned_kg10 += kg10
        entry.planned_kg20 += kg20
        entry.planned_kg40 += kg40
        entry.planned_kg50 += kg50
        // Obs: pega a observação da linha que tiver, sem sobrescrever
        if (!entry.obs && r[17]) entry.obs = String(r[17]).trim()
      } else {
        // Primeira linha deste VDA → cria entrada, guarda rota/motorista/cap
        const entry = {
          vda:            vdaKey,
          rota:           r[0] ? String(r[0]).trim() : null,
          motorista_name: r[2] ? String(r[2]).trim() : null,
          capacity_kg:    numOf(r[3]),
          planned_kg3:    kg3,
          planned_kg5:    kg5,
          planned_kg10:   kg10,
          planned_kg20:   kg20,
          planned_kg40:   kg40,
          planned_kg50:   kg50,
          obs:            r[17] ? String(r[17]).trim() : null,
        }
        vdaMap.set(vdaKey, entry)
        vdaOrder.push(vdaKey)
      }
    }

    // Filtra VDAs sem nenhuma quantidade e sem capacidade (linhas em branco/cabeçalho)
    const vehicles = []
    let seq = 0
    for (const key of vdaOrder) {
      const v = vdaMap.get(key)
      const totalSacos = v.planned_kg3 + v.planned_kg5 + v.planned_kg10 + v.planned_kg20 + v.planned_kg40 + v.planned_kg50
      if (totalSacos === 0 && v.capacity_kg === 0) continue
      seq++
      vehicles.push({ sequence: seq, ...v })
    }

    if (vehicles.length === 0) throw Object.assign(new Error("Nenhum veiculo encontrado"), { code: 422 })
    return { date: operationDate, sheetName, vehicles }
  }

  // ─── POST /plans/parse-excel — parseia xlsx, retorna preview SEM salvar ───
  fastify.post("/plans/parse-excel", async (req, reply) => {
    try {
      const data = await req.file()
      if (!data) return reply.code(400).send({ error: "Campo 'file' ausente" })
      const buffer = await data.toBuffer()
      const dateParam = req.body?.date || null
      const result = parseXlsxBuffer(buffer, dateParam)
      return reply.code(200).send(result)
    } catch (err) {
      fastify.log.error({ err }, "parse-excel error")
      return reply.code(err.code || 422).send({ error: err.message, sheets_available: err.sheets_available })
    }
  })

  // ─── Importação via JSON (fluxo original + fluxo do painel-central) ────────
  fastify.post("/plans/import", async (req, reply) => {
    const plan = req.body
    if (!plan?.vehicles?.length) {
      return reply.code(400).send({ error: "JSON invalido — sem veiculos" })
    }
    const client = await fastify.db.pool.connect()
    try {
      await client.query("BEGIN")

      // Suporta dois formatos de data:
      //   { date: "YYYY-MM-DD" }  ← painel-central (novo)
      //   { source_date: "DD/MM/YYYY" }  ← formato legado
      let operationDate = plan.date
      if (!operationDate) {
        const dateMatch = plan.source_date?.match(/(\d{2})\/(\d{2})\/(\d{4})/)
        if (!dateMatch) return reply.code(400).send({ error: "Data nao encontrada no JSON" })
        operationDate = `${dateMatch[3]}-${dateMatch[2]}-${dateMatch[1]}`
      }
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
        // Detecta se é formato novo (planned_kg*) ou legado (volumes[])
        const isNewFormat = "planned_kg5" in v || "planned_kg10" in v

        const { rows: [veh] } = await client.query(
          `INSERT INTO vehicles (daily_plan_id, sequence, rota, vda, vehicle_type, capacity_kg, motorista_name, equipe, state)
           VALUES ($1, $2, $3, $4, $5::vehicle_type_enum, $6, $7, $8, 'aguardando_carga') RETURNING id`,
          [planId, v.sequence, v.rota, v.vda,
           (v.vehicle_type || "unknown") === "3/4" ? "3/4" : (v.vehicle_type || "unknown"),
           v.capacity_kg || 0, v.motorista_name || v.motorista || null, v.equipe || []]
        )
        vehicleIds.push({ vda: v.vda, id: veh.id })

        if (isNewFormat) {
          // Formato painel-central: um volume "PRE-VENDA" com os SKUs direto
          const peso = (v.planned_kg3||0)*3 + (v.planned_kg5||0)*5 + (v.planned_kg10||0)*10
                     + (v.planned_kg20||0)*20 + (v.planned_kg40||0)*40 + (v.planned_kg50||0)*50
          await client.query(
            `INSERT INTO vehicle_volumes
               (vehicle_id, volume_type, return_rule, oc_pending,
                planned_kg3, planned_kg5, planned_kg10, planned_kg20, planned_kg40, planned_kg50, planned_peso_kg, obs)
             VALUES ($1, 'PRE-VENDA'::volume_type_enum, 'none'::return_rule_enum, false,
                     $2, $3, $4, $5, $6, $7, $8, $9)`,
            [veh.id, v.planned_kg3||0, v.planned_kg5||0, v.planned_kg10||0,
             v.planned_kg20||0, v.planned_kg40||0, v.planned_kg50||0, peso, v.obs||null]
          )
        } else {
          // Formato legado com volumes[]
          for (const vol of (v.volumes || [])) {
            await client.query(
              `INSERT INTO vehicle_volumes
                 (vehicle_id, volume_type, top_sankhya, oc_exclusive, return_rule, oc_number, oc_pending,
                  planned_kg3, planned_kg5, planned_kg10, planned_kg20, planned_kg40, planned_kg50, planned_peso_kg, obs)
               VALUES ($1, $2::volume_type_enum, $3, $4, $5::return_rule_enum, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15)`,
              [veh.id, vol.type, vol.top_sankhya, vol.oc_exclusive, vol.return_rule, vol.oc_number, vol.oc_pending,
               vol.skus?.kg3||0, vol.skus?.kg5||0, vol.skus?.kg10||0, vol.skus?.kg20||0,
               vol.skus?.kg40||0, vol.skus?.kg50||0, vol.peso_kg||0, vol.obs||null]
            )
          }
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

  // ─── Importação da Programação Excel ───────────────────────────────────────
  fastify.post("/plans/import-programacao", async (req, reply) => {
    try {
      const data = await req.file()
      if (!data) return reply.code(400).send({ error: "Arquivo nao encontrado no campo 'file'" })

      const buffer = await data.toBuffer()
      const wb = XLSX.read(buffer, { type: "buffer", raw: true })

      const targetDateStr = req.body?.date
      let operationDate
      if (targetDateStr) {
        if (/^\d{4}-\d{2}-\d{2}$/.test(targetDateStr)) {
          operationDate = targetDateStr
        } else {
          const m = targetDateStr.match(/^(\d{2})\.(\d{2})\.(\d{4})$/)
          if (m) operationDate = `${m[3]}-${m[2]}-${m[1]}`
        }
      }
      if (!operationDate) {
        const now = new Date(Date.now() - 3 * 60 * 60 * 1000)
        const [y, m, d] = now.toISOString().split("T")[0].split("-")
        operationDate = `${y}-${m}-${d}`
      }

      const [y, m, d] = operationDate.split("-")
      const sheetName = `${d}.${m}.${y}`
      const ws = wb.Sheets[sheetName]
      if (!ws) {
        const available = wb.SheetNames.join(", ")
        return reply.code(404).send({ error: `Sheet '${sheetName}' nao encontrada no arquivo`, sheets_available: available })
      }

      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: false })

      let headerRow = -1
      for (let i = 0; i < Math.min(raw.length, 10); i++) {
        const row = raw[i]
        if (row && row.some(c => typeof c === "string" && c.toUpperCase().includes("VDA"))) {
          headerRow = i
          break
        }
      }
      if (headerRow === -1) {
        return reply.code(422).send({ error: "Cabecalho com 'VDA' nao encontrado na sheet" })
      }

      const numOf = (v) => {
        if (v === null || v === undefined || v === "") return 0
        const n = Number(v)
        return isNaN(n) ? 0 : Math.round(n)
      }

      const toKey = (v) => {
        if (v === null || v === undefined) return null
        const s = String(v).trim().toUpperCase()
        if (s.includes("ACCELO") || s.includes("ACELO")) return "ACCELO"
        const digits = s.replace(/\D/g, "")
        if (!digits) return null
        return String(parseInt(digits, 10))
      }

      const vehicles = []
      let seq = 0
      for (let i = headerRow + 1; i < raw.length; i++) {
        const r = raw[i]
        if (!r || r.every(c => c === null)) continue
        const vdaKey = toKey(r[1])
        if (!vdaKey) continue
        const rota       = r[0] ? String(r[0]).trim() : null
        const capacidade = numOf(r[3])
        const kg5        = numOf(r[11])
        const kg10       = numOf(r[12])
        const kg20       = numOf(r[13])
        const kg40       = numOf(r[14])
        const obs        = r[17] ? String(r[17]).trim() : null
        if (kg5 === 0 && kg10 === 0 && kg20 === 0 && kg40 === 0 && capacidade === 0) continue
        seq++
        vehicles.push({ seq, vda: vdaKey, rota, capacity_kg: capacidade, kg5, kg10, kg20, kg40, obs })
      }

      if (vehicles.length === 0) {
        return reply.code(422).send({ error: "Nenhum veiculo com dados encontrado na sheet" })
      }

      const { rows: [plan] } = await fastify.db.query(
        `SELECT id FROM daily_plans WHERE operation_date = $1`, [operationDate]
      )
      if (!plan) {
        return reply.code(404).send({
          error: `Plano para ${operationDate} nao encontrado. Importe o plano base primeiro via /api/plans/import`,
          vehicles_parsed: vehicles
        })
      }
      const planId = plan.id

      const client = await fastify.db.pool.connect()
      const results = []
      try {
        await client.query("BEGIN")
        for (const v of vehicles) {
          const { rows: [veh] } = await client.query(
            `SELECT id FROM vehicles WHERE daily_plan_id = $1 AND vda = $2`, [planId, v.vda]
          )
          if (!veh) { results.push({ vda: v.vda, status: "nao_encontrado_no_plano" }); continue }
          const { rows: [vol] } = await client.query(
            `SELECT id FROM vehicle_volumes WHERE vehicle_id = $1 LIMIT 1`, [veh.id]
          )
          if (vol) {
            await client.query(
              `UPDATE vehicle_volumes SET planned_kg5=$1, planned_kg10=$2, planned_kg20=$3, planned_kg40=$4, obs=COALESCE($5,obs) WHERE id=$6`,
              [v.kg5, v.kg10, v.kg20, v.kg40, v.obs, vol.id]
            )
            results.push({ vda: v.vda, status: "atualizado", kg5: v.kg5, kg10: v.kg10, kg20: v.kg20, kg40: v.kg40 })
          } else {
            await client.query(
              `INSERT INTO vehicle_volumes (vehicle_id, volume_type, planned_kg5, planned_kg10, planned_kg20, planned_kg40, obs) VALUES ($1,'frio'::volume_type_enum,$2,$3,$4,$5,$6)`,
              [veh.id, v.kg5, v.kg10, v.kg20, v.kg40, v.obs]
            )
            results.push({ vda: v.vda, status: "criado", kg5: v.kg5, kg10: v.kg10, kg20: v.kg20, kg40: v.kg40 })
          }
        }
        await client.query("COMMIT")
      } catch (err) {
        await client.query("ROLLBACK")
        fastify.log.error({ err }, "Erro ao salvar programacao")
        return reply.code(500).send({ error: err.message })
      } finally {
        client.release()
      }

      return reply.code(200).send({
        operationDate, sheetName,
        vehicles_parsed: vehicles.length,
        updated:  results.filter(r => r.status === "atualizado").length,
        created:  results.filter(r => r.status === "criado").length,
        not_found_in_plan: results.filter(r => r.status === "nao_encontrado_no_plano").length,
        results
      })
    } catch (err) {
      fastify.log.error({ err }, "Erro ao processar arquivo de programacao")
      return reply.code(500).send({ error: err.message })
    }
  })

  // ─── GET /plans/:date ──────────────────────────────────────────────────────
  fastify.get("/plans/:date", async (req, reply) => {
    const { rows: [plan] } = await fastify.db.query(
      `SELECT id, operation_date, sheet_name, total_vehicles, imported_at FROM daily_plans WHERE operation_date = $1`,
      [req.params.date]
    )
    if (!plan) return reply.code(404).send({ error: "Plano nao encontrado" })
    return plan
  })

  // ─── GET /plans/:date/vehicles (inclui SKUs planejados) ───────────────────
  fastify.get("/plans/:date/vehicles", async (req, reply) => {
    const { rows } = await fastify.db.query(
      `SELECT
         v.id, v.sequence, v.rota, v.vda, v.vehicle_type, v.capacity_kg,
         v.motorista_name, v.equipe, v.state,
         v.paletes_realizados, v.estrados_realizados, v.lotes, v.km_saida, v.km_retorno,
         v.carga_inicio_ts, v.carga_fim_ts, v.liberacao_asistente_ts, v.saida_portaria_ts,
         v.retorno_fisico_ts, v.finalizado_ts, v.duracao_carga_min, v.duracao_liberacao_min,
         v.liberacao_atrasada,
         COALESCE(SUM(vv.planned_kg3),  0)::int AS planned_kg3,
         COALESCE(SUM(vv.planned_kg5),  0)::int AS planned_kg5,
         COALESCE(SUM(vv.planned_kg10), 0)::int AS planned_kg10,
         COALESCE(SUM(vv.planned_kg20), 0)::int AS planned_kg20,
         COALESCE(SUM(vv.planned_kg40), 0)::int AS planned_kg40,
         COALESCE(SUM(vv.planned_kg50), 0)::int AS planned_kg50
       FROM vehicles v
       JOIN daily_plans dp ON dp.id = v.daily_plan_id
       LEFT JOIN vehicle_volumes vv ON vv.vehicle_id = v.id
       WHERE dp.operation_date = $1
       GROUP BY v.id
       ORDER BY v.sequence`,
      [req.params.date]
    )
    return rows
  })
}
