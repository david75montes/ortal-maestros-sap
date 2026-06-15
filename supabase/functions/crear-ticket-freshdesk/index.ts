const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const { folio, solicitante_nombre, solicitante_email, planillas, attachments } = await req.json();

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
Esta solicitud fue aprobada en el portal NEXUS y requiere carga en SAP.<br>
Los archivos Excel SAP están adjuntos a este ticket.
    `.trim();

    const credentials = btoa(`${apiKey}:X`);

    const formData = new FormData();
    formData.append("subject", `[NEXUS] Solicitud ${folio}`);
    formData.append("description", description);
    formData.append("email", solicitante_email);
    formData.append("priority", "2");
    formData.append("status", "2");
    formData.append("tags[]", "nexus");
    formData.append("tags[]", "datos-maestros");

    if (attachments && attachments.length > 0) {
      for (const att of (attachments as { base64: string; filename: string }[])) {
        const binStr = atob(att.base64);
        const bytes = new Uint8Array(binStr.length);
        for (let i = 0; i < binStr.length; i++) bytes[i] = binStr.charCodeAt(i);
        const blob = new Blob([bytes], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        formData.append("attachments[]", blob, att.filename);
      }
    }

    const response = await fetch(`https://${domain}/api/v2/tickets`, {
      method: "POST",
      headers: { "Authorization": `Basic ${credentials}` },
      body: formData,
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
