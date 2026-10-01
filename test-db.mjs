import pg from "pg"
const { Pool } = pg
const pool = new Pool({ connectionString: "postgresql://postgres.blmpymcmnhmjvibfilem:ajm120850vame270289@aws-0-sa-east-1.pooler.supabase.com:6543/postgres" })
const { rows } = await pool.query("SELECT COUNT(*) as total FROM actors")
console.log("Conexao OK — Atores cadastrados:", rows[0].total)
await pool.end()
