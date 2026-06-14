const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const { folio, solicitante_nombre, solicitante_email, planillas } = await req.json();

    const apiKey = Deno.env.get("FRESHDESK_API_KEY");
    const domain = "ddmm.freshdesk.com";

    const planillasDesc = (planillas || [])
      .map((p: { nombre: string; filas: unknown[] }) => `• ${p.nombre}: ${p.filas.length} filas`)
      .join("<br>");

    const description = `
<b>Solicitud NEXUS aprobada: ${folio}</b><br><br>
<b>Solicitante:</b> ${solicitante_nombre} (${solicitante_email})<br><br>
<b>Planillas a procesar:</b><br>
${planillasDesc}<br><br>
Esta solicitud fue aprobada en el portal NEXUS y requiere carga en SAP.
    `.trim();

    const ticketData = {
      subject: `[NEXUS] Solicitud ${folio}`,
      description,
      email: solicitante_email,
      priority: 2,
      status: 2,
      tags: ["nexus", "datos-maestros"],
    };

    const credentials = btoa(`${apiKey}:X`);

    const response = await fetch(`https://${domain}/api/v2/tickets`, {
      method: "POST",
      headers: {
        "Authorization": `Basic ${credentials}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(ticketData),
    });

    if (!response.ok) {
      const err = await response.text();
      return new Response(JSON.stringify({ error: err }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const ticket = await response.json();
    return new Response(JSON.stringify({ ok: true, ticket_id: ticket.id }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
