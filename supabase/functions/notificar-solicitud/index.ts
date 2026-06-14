import nodemailer from "npm:nodemailer@6";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const { folio, estado, solicitante_nombre, solicitante_email, planillas, motivo_rechazo } = await req.json();

    if (!solicitante_email) {
      return new Response(JSON.stringify({ error: "Sin email del solicitante" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const transporter = nodemailer.createTransport({
      host: "smtp.gmail.com",
      port: 587,
      secure: false,
      auth: {
        user: Deno.env.get("GMAIL_USER"),
        pass: Deno.env.get("GMAIL_APP_PASSWORD"),
      },
    });

    const aprobada = estado === "Aprobada";
    const subject = `[NEXUS] Solicitud ${folio} — ${aprobada ? "Aprobada ✓" : "Rechazada ✗"}`;

    const planillasTexto = (planillas || [])
      .map((p: { nombre: string; filas: unknown[] }) => `  • ${p.nombre}: ${p.filas.length} filas`)
      .join("\n");

    const cuerpo = aprobada
      ? `Hola ${solicitante_nombre},

Tu solicitud ${folio} ha sido APROBADA por el equipo de Datos Maestros.

Planillas incluidas:
${planillasTexto}

El equipo de DDMM procesará los cambios en SAP próximamente y actualizará el estado a "Aplicada".

Puedes seguir el avance en el portal NEXUS → Solicitudes.

Saludos,
NEXUS · Arcoprime`
      : `Hola ${solicitante_nombre},

Tu solicitud ${folio} ha sido RECHAZADA por el equipo de Datos Maestros.

${motivo_rechazo ? `Motivo: ${motivo_rechazo}\n` : ""}
Si tienes consultas, contacta al equipo de Datos Maestros respondiendo este correo.

Saludos,
NEXUS · Arcoprime`;

    await transporter.sendMail({
      from: `"NEXUS · Arcoprime" <${Deno.env.get("GMAIL_USER")}>`,
      to: solicitante_email,
      subject,
      text: cuerpo,
    });

    return new Response(JSON.stringify({ ok: true }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
