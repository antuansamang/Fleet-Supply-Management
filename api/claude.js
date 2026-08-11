// Función serverless de Vercel. Corre en el servidor, nunca en el navegador del usuario.
// Recibe { messages, max_tokens } desde el frontend y llama a la API de Anthropic
// usando la clave guardada en la variable de entorno ANTHROPIC_API_KEY (nunca expuesta al cliente).
//
// Vercel rechaza (HTTP 413) cualquier cuerpo de solicitud a esta función que supere ~4.5MB.
// Los PDFs escaneados grandes no caben ahí una vez codificados en base64, así que el
// frontend los sube primero a un bucket temporal de Supabase Storage y solo manda una
// referencia (bucket + ruta). Esta función descarga el archivo real desde Storage y arma
// el bloque de contenido base64 que espera la API de Anthropic antes de reenviarlo.

async function resolveBlock(block) {
  if (!block || !block.source || block.source.type !== "storage_ref") return block;
  const { bucket, path, media_type } = block.source;
  const supabaseUrl = process.env.VITE_SUPABASE_URL;
  if (!supabaseUrl) {
    throw new Error("Falta VITE_SUPABASE_URL en el servidor para leer el archivo temporal de Storage.");
  }
  const objectUrl = supabaseUrl.replace(/\/$/, "") + "/storage/v1/object/public/" + bucket + "/" + path;
  const fileRes = await fetch(objectUrl);
  if (!fileRes.ok) {
    throw new Error("No se pudo leer el archivo temporal en Storage (HTTP " + fileRes.status + ").");
  }
  const buf = Buffer.from(await fileRes.arrayBuffer());
  return { type: block.type, source: { type: "base64", media_type, data: buf.toString("base64") } };
}

async function resolveMessages(messages) {
  return Promise.all(
    (messages || []).map(async (m) => {
      if (!Array.isArray(m.content)) return m;
      return { ...m, content: await Promise.all(m.content.map(resolveBlock)) };
    })
  );
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: { message: "Método no permitido" } });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: { message: "ANTHROPIC_API_KEY no está configurada en el servidor." } });
    return;
  }

  try {
    const { messages, max_tokens } = req.body || {};
    if (!messages) {
      res.status(400).json({ error: { message: "Falta 'messages' en la solicitud." } });
      return;
    }

    const resolvedMessages = await resolveMessages(messages);

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: max_tokens || 4000,
        messages: resolvedMessages,
      }),
    });

    const data = await response.json();
    res.status(response.status).json(data);
  } catch (err) {
    res.status(500).json({ error: { message: "Error interno del proxy: " + (err && err.message) } });
  }
}
