import XLSX from "xlsx"

export default async function plansRoutes(fastify) {

  // ─── Helper: parseia buffer xlsx → preview de veículos ─────────────────────
  function parseXlsxBuffer(buffer, targetDate) {
    const wb = XLSX.read(buffer, { type: "buffer", raw: true })

    // Determinar data de operação
    let operationDate = targetDate
    if (!operationDate) {
      const now = new Date(Date.now() - 4 * 60 * 60 * 1000)  // Manaus UTC-4
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

    // Normaliza o tipo de volume da coluna J (col[9])
    // Possíveis valores: PRE-VENDA, TROCAS, MANIFESTO, CONSIGNADA, BONIFICAÇÃO, SALDO
    // toVolumeType com cache — normalize roda no máximo 1x por valor único
    const _vtCache = new Map()
    const toVolumeType = (v) => {
      if (!v) return null
      if (_vtCache.has(v)) return _vtCache.get(v)
      const s = String(v).trim().toUpperCase()
        .normalize("NFD").replace(/[̀-ͯ]/g, "")
      let result
      if (s.includes("PRE") && s.includes("VENDA")) result = "PRE-VENDA"
      else if (s.includes("TROCA"))   result = "TROCAS"
      else if (s.includes("MANIFES")) result = "MANIFESTO"
      else if (s.includes("CONSIG"))  result = "CONSIGNADA"
      else if (s.includes("BONIF"))   result = "BONIFICAÇÃO"
      else if (s.includes("SALDO"))   result = "SALDO"
      else result = s
      _vtCache.set(v, result)
      return result
    }

    // Agrupa linhas por viagem (VDA + ocorrência).
    // O mesmo VDA pode aparecer duas vezes na planilha quando há recarga (segunda viagem).
    // Detectamos isso rastreando qual VDA estava ativo: quando o VDA muda de volta para um
    // já visto após ter passado por outro VDA, é uma nova entrada — não acumulamos.
    //
    // Chave interna: "${vdaKey}#${N}" onde N começa em 0 (a exportada é só vda + recarga flag).
    // Map preserva ordem de inserção (ES2015+).
    const vdaMap   = new Map()   // chave composta → entry
    const vdaCount = new Map()   // vdaKey → quantas entradas já criadas
    let activeKey  = null        // chave composta da entrada corrente

    for (let i = headerRow + 1; i < raw.length; i++) {
      const r = raw[i]
      if (!r || r.every(c => c === null)) continue
      const vdaKey = toKey(r[1])
      if (!vdaKey) continue

      const volumeType = toVolumeType(r[9])

      const kg3  = numOf(r[10]), kg5  = numOf(r[11]), kg10 = numOf(r[12])
      const kg20 = numOf(r[13]), kg40 = numOf(r[14]), kg50 = numOf(r[15])
      const pesoKg = kg3*3 + kg5*5 + kg10*10 + kg20*20 + kg40*40 + kg50*50

      if (!volumeType && pesoKg === 0) continue

      const rawOc = r[7] ? String(r[7]).trim() : null
      const ocNumber = rawOc && rawOc !== "" && rawOc !== "0" ? rawOc : null
      const ocPending = ocNumber === null && volumeType !== "SALDO" && volumeType !== null
      const obs = r[17] ? String(r[17]).trim() : null

      // Determina se esta linha pertence à entrada ativa ou inicia uma nova
      const activeVda = activeKey ? activeKey.split("#")[0] : null
      if (activeVda !== vdaKey) {
        // VDA mudou — verifica se é um VDA novo ou uma nova ocorrência de um já visto
        const n = vdaCount.get(vdaKey) ?? 0
        const compositeKey = `${vdaKey}#${n}`
        vdaCount.set(vdaKey, n + 1)
        activeKey = compositeKey

        const entry = {
          vda:            vdaKey,
          recarga:        n > 0,                           // true a partir da 2ª viagem
          rota:           r[0] ? String(r[0]).trim() : null,
          motorista_name: r[2] ? String(r[2]).trim() : null,
          capacity_kg:    numOf(r[3]),
          planned_kg3:  kg3,  planned_kg5:  kg5,  planned_kg10: kg10,
          planned_kg20: kg20, planned_kg40: kg40, planned_kg50: kg50,
          volumes: [],
        }
        if (volumeType || pesoKg > 0) {
          entry.volumes.push({ type: volumeType, oc_number: ocNumber, oc_pending: ocPending,
            kg3, kg5, kg10, kg20, kg40, kg50, peso_kg: pesoKg, obs })
        }
        vdaMap.set(compositeKey, entry)
      } else {
        // Mesma entrada ativa — acumula (ex: linha PRE-VENDA + linha TROCAS do mesmo bloco)
        const entry = vdaMap.get(activeKey)
        entry.planned_kg3  += kg3;  entry.planned_kg5  += kg5;  entry.planned_kg10 += kg10
        entry.planned_kg20 += kg20; entry.planned_kg40 += kg40; entry.planned_kg50 += kg50
        if (volumeType || pesoKg > 0) {
          entry.volumes.push({ type: volumeType, oc_number: ocNumber, oc_pending: ocPending,
            kg3, kg5, kg10, kg20, kg40, kg50, peso_kg: pesoKg, obs })
        }
      }
    }

    // Monta array final — Map já está em ordem de inserção
    const vehicles = []
    let seq = 0
    for (const [, v] of vdaMap) {
      const totalSacos = v.planned_kg3 + v.planned_kg5 + v.planned_kg10 + v.planned_kg20 + v.planned_kg40 + v.planned_kg50
      if (totalSacos === 0 && v.capacity_kg === 0 && v.volumes.length === 0) continue
      seq++
      // Campo de conveniência: primeiro OC encontrado (compatibilidade com tabela do painel-central)
      const firstOc = v.volumes.find(vol => vol.oc_number)
      const hasPending = v.volumes.some(vol => vol.oc_pending)
      vehicles.push({
        sequence: seq,
        vda: v.vda,
        rota: v.rota,
        motorista_name: v.motorista_name,
        capacity_kg: v.capacity_kg,
        planned_kg3:  v.planned_kg3,
        planned_kg5:  v.planned_kg5,
        planned_kg10: v.planned_kg10,
        planned_kg20: v.planned_kg20,
        planned_kg40: v.planned_kg40,
        planned_kg50: v.planned_kg50,
        oc_number:  firstOc ? firstOc.oc_number : null,
        oc_pending: hasPending,
        volumes: v.volumes,   // detalhamento completo por tipo
      })
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
      // Bulk insert de veículos — 1 query para todos (O(1) round-trips)
      const vSeqs = [], vRotas = [], vVdas = [], vTypes = [], vCaps = [], vMots = [], vEquipes = []
      for (const v of plan.vehicles) {
        vSeqs.push(v.sequence)
        vRotas.push(v.rota)
        vVdas.push(v.vda)
        vTypes.push((v.vehicle_type || "unknown") === "3/4" ? "3/4" : (v.vehicle_type || "unknown"))
        vCaps.push(v.capacity_kg || 0)
        vMots.push(v.motorista_name || v.motorista || null)
        vEquipes.push(v.equipe || [])
      }
      const { rows: insertedVehicles } = await client.query(
        `INSERT INTO vehicles (daily_plan_id, sequence, rota, vda, vehicle_type, capacity_kg, motorista_name, equipe, state)
         SELECT $1, unnest($2::int[]), unnest($3::text[]), unnest($4::text[]),
                unnest($5::vehicle_type_enum[]), unnest($6::int[]), unnest($7::text[]),
                unnest($8::text[][]), 'aguardando_carga'
         RETURNING id, vda, sequence`,
        [planId, vSeqs, vRotas, vVdas, vTypes, vCaps, vMots, vEquipes]
      )
      // Correlaciona por sequence — único por plano, mesmo VDA podendo aparecer 2×
      const vehicleIdBySeq = new Map(insertedVehicles.map(r => [r.sequence, r.id]))
      const vehicleIds = insertedVehicles.map(r => ({ vda: r.vda, id: r.id }))

      // Bulk insert de volumes — 1 query para todos os volumes de todos os veículos
      const volVehicleIds = [], volTypes = [], volTopS = [], volOcExcl = [], volRetRule = []
      const volOcNum = [], volOcPend = []
      const volKg3 = [], volKg5 = [], volKg10 = [], volKg20 = [], volKg40 = [], volKg50 = []
      const volPeso = [], volObs = []

      for (const v of plan.vehicles) {
        const vehicleId = vehicleIdBySeq.get(v.sequence)
        if (!vehicleId) continue
        const isNewFormat = "planned_kg5" in v || "planned_kg10" in v

        if (isNewFormat) {
          const peso = (v.planned_kg3||0)*3 + (v.planned_kg5||0)*5 + (v.planned_kg10||0)*10
                     + (v.planned_kg20||0)*20 + (v.planned_kg40||0)*40 + (v.planned_kg50||0)*50
          volVehicleIds.push(vehicleId); volTypes.push("PRE-VENDA"); volTopS.push(null)
          volOcExcl.push(null); volRetRule.push("none"); volOcNum.push(null); volOcPend.push(false)
          volKg3.push(v.planned_kg3||0); volKg5.push(v.planned_kg5||0); volKg10.push(v.planned_kg10||0)
          volKg20.push(v.planned_kg20||0); volKg40.push(v.planned_kg40||0); volKg50.push(v.planned_kg50||0)
          volPeso.push(peso); volObs.push(v.obs||null)
        } else {
          for (const vol of (v.volumes || [])) {
            volVehicleIds.push(vehicleId); volTypes.push(vol.type); volTopS.push(vol.top_sankhya||null)
            volOcExcl.push(vol.oc_exclusive||null); volRetRule.push(vol.return_rule||"none")
            volOcNum.push(vol.oc_number||null); volOcPend.push(vol.oc_pending||false)
            volKg3.push(vol.skus?.kg3||0); volKg5.push(vol.skus?.kg5||0); volKg10.push(vol.skus?.kg10||0)
            volKg20.push(vol.skus?.kg20||0); volKg40.push(vol.skus?.kg40||0); volKg50.push(vol.skus?.kg50||0)
            volPeso.push(vol.peso_kg||0); volObs.push(vol.obs||null)
          }
        }
      }

      if (volVehicleIds.length > 0) {
        await client.query(
          `INSERT INTO vehicle_volumes
             (vehicle_id, volume_type, top_sankhya, oc_exclusive, return_rule, oc_number, oc_pending,
              planned_kg3, planned_kg5, planned_kg10, planned_kg20, planned_kg40, planned_kg50, planned_peso_kg, obs)
           SELECT unnest($1::int[]), unnest($2::volume_type_enum[]), unnest($3::text[]), unnest($4::text[]),
                  unnest($5::return_rule_enum[]), unnest($6::text[]), unnest($7::bool[]),
                  unnest($8::int[]), unnest($9::int[]), unnest($10::int[]), unnest($11::int[]),
                  unnest($12::int[]), unnest($13::int[]), unnest($14::int[]), unnest($15::text[])`,
          [volVehicleIds, volTypes, volTopS, volOcExcl, volRetRule, volOcNum, volOcPend,
           volKg3, volKg5, volKg10, volKg20, volKg40, volKg50, volPeso, volObs]
        )
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
  // F) Parser unificado — reutiliza parseXlsxBuffer para sheet name / header / numOf / toKey
  fastify.post("/plans/import-programacao", async (req, reply) => {
    try {
      const data = await req.file()
      if (!data) return reply.code(400).send({ error: "Arquivo nao encontrado no campo 'file'" })

      const buffer = await data.toBuffer()

      // Normaliza targetDate para YYYY-MM-DD
      const targetDateStr = req.body?.date
      let targetDate = null
      if (targetDateStr) {
        if (/^\d{4}-\d{2}-\d{2}$/.test(targetDateStr)) {
          targetDate = targetDateStr
        } else {
          const mDate = targetDateStr.match(/^(\d{2})\.(\d{2})\.(\d{4})$/)
          if (mDate) targetDate = `${mDate[3]}-${mDate[2]}-${mDate[1]}`
        }
      }

      // Reutiliza parseXlsxBuffer — mesmos helpers, mesma lógica de sheet
      let parsed
      try {
        parsed = parseXlsxBuffer(buffer, targetDate)
      } catch (err) {
        return reply.code(err.code || 422).send({ error: err.message, sheets_available: err.sheets_available })
      }

      const operationDate = parsed.date
      const sheetName     = parsed.sheetName

      // Extrai apenas os campos relevantes para programação (kg5/10/20/40)
      // v.sequence é único por entry (mesmo VDA pode aparecer 2× com sequences distintos)
      const vehicles = parsed.vehicles
        .map((v) => {
          const kg5  = v.planned_kg5  || 0
          const kg10 = v.planned_kg10 || 0
          const kg20 = v.planned_kg20 || 0
          const kg40 = v.planned_kg40 || 0
          const obs  = v.volumes?.[0]?.obs || null
          if (kg5 === 0 && kg10 === 0 && kg20 === 0 && kg40 === 0 && v.capacity_kg === 0) return null
          return { seq: v.sequence, vda: v.vda, rota: v.rota, capacity_kg: v.capacity_kg, kg5, kg10, kg20, kg40, obs }
        })
        .filter(Boolean)

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

      // ── E) Bulk lookup + update/insert — O(3) queries independente de N ──────
      // Usa sequence como chave — único por plano, mesmo VDA podendo aparecer 2× (carga + recarga)
      const seqList  = vehicles.map(v => v.seq)
      const { rows: existingVehicles } = await fastify.db.query(
        `SELECT id, vda, sequence FROM vehicles WHERE daily_plan_id = $1 AND sequence = ANY($2::int[])`,
        [planId, seqList]
      )
      const vehicleIdBySeq = new Map(existingVehicles.map(r => [r.sequence, r.id]))
      const foundSeqs = new Set(existingVehicles.map(r => r.sequence))

      // Separa veículos encontrados dos não encontrados
      const notFound  = vehicles.filter(v => !foundSeqs.has(v.seq))
      const toProcess = vehicles.filter(v =>  foundSeqs.has(v.seq))

      const results = notFound.map(v => ({ vda: v.vda, seq: v.seq, status: "nao_encontrado_no_plano" }))

      if (toProcess.length > 0) {
        const vehicleIds = toProcess.map(v => vehicleIdBySeq.get(v.seq))

        // Busca o primeiro volume de cada veículo em bulk
        const { rows: existingVols } = await fastify.db.query(
          `SELECT DISTINCT ON (vehicle_id) id, vehicle_id
           FROM vehicle_volumes WHERE vehicle_id = ANY($1::int[])
           ORDER BY vehicle_id, id`,
          [vehicleIds]
        )
        const volIdByVehicleId = new Map(existingVols.map(r => [r.vehicle_id, r.id]))

        const toUpdate = toProcess.filter(v => volIdByVehicleId.has(vehicleIdBySeq.get(v.seq)))
        const toCreate = toProcess.filter(v => !volIdByVehicleId.has(vehicleIdBySeq.get(v.seq)))

        const client = await fastify.db.pool.connect()
        try {
          await client.query("BEGIN")

          // UPDATE em bulk com UNNEST
          if (toUpdate.length > 0) {
            const upIds  = toUpdate.map(v => volIdByVehicleId.get(vehicleIdBySeq.get(v.seq)))
            const upKg5  = toUpdate.map(v => v.kg5)
            const upKg10 = toUpdate.map(v => v.kg10)
            const upKg20 = toUpdate.map(v => v.kg20)
            const upKg40 = toUpdate.map(v => v.kg40)
            const upObs  = toUpdate.map(v => v.obs || null)
            await client.query(
              `UPDATE vehicle_volumes AS vv SET
                 planned_kg5  = u.kg5,
                 planned_kg10 = u.kg10,
                 planned_kg20 = u.kg20,
                 planned_kg40 = u.kg40,
                 obs = COALESCE(u.obs, vv.obs)
               FROM (SELECT unnest($1::int[]) AS id, unnest($2::int[]) AS kg5,
                            unnest($3::int[]) AS kg10, unnest($4::int[]) AS kg20,
                            unnest($5::int[]) AS kg40, unnest($6::text[]) AS obs) AS u
               WHERE vv.id = u.id`,
              [upIds, upKg5, upKg10, upKg20, upKg40, upObs]
            )
            for (const v of toUpdate)
              results.push({ vda: v.vda, seq: v.seq, status: "atualizado", kg5: v.kg5, kg10: v.kg10, kg20: v.kg20, kg40: v.kg40 })
          }

          // INSERT em bulk com UNNEST
          if (toCreate.length > 0) {
            const cIds  = toCreate.map(v => vehicleIdBySeq.get(v.seq))
            const cKg5  = toCreate.map(v => v.kg5)
            const cKg10 = toCreate.map(v => v.kg10)
            const cKg20 = toCreate.map(v => v.kg20)
            const cKg40 = toCreate.map(v => v.kg40)
            const cObs  = toCreate.map(v => v.obs || null)
            await client.query(
              `INSERT INTO vehicle_volumes (vehicle_id, volume_type, planned_kg5, planned_kg10, planned_kg20, planned_kg40, obs)
               SELECT unnest($1::int[]), 'frio'::volume_type_enum,
                      unnest($2::int[]), unnest($3::int[]), unnest($4::int[]),
                      unnest($5::int[]), unnest($6::text[])`,
              [cIds, cKg5, cKg10, cKg20, cKg40, cObs]
            )
            for (const v of toCreate)
              results.push({ vda: v.vda, seq: v.seq, status: "criado", kg5: v.kg5, kg10: v.kg10, kg20: v.kg20, kg40: v.kg40 })
          }

          await client.query("COMMIT")
        } catch (err) {
          await client.query("ROLLBACK")
          fastify.log.error({ err }, "Erro ao salvar programacao")
          return reply.code(500).send({ error: err.message })
        } finally {
          client.release()
        }
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
