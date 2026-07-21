// Función serverless de Vercel. Corre en el servidor, nunca en el navegador del usuario.
// Recibe { messages, max_tokens } desde el frontend y llama a la API de Anthropic
// usando la clave guardada en la variable de entorno ANTHROPIC_API_KEY (nunca expuesta al cliente).

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
        messages,
      }),
    });

    const data = await response.json();
    res.status(response.status).json(data);
  } catch (err) {
    res.status(500).json({ error: { message: "Error interno del proxy: " + (err && err.message) } });
  }
}
