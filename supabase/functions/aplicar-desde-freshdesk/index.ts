import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import nodemailer from "npm:nodemailer@6";

Deno.serve(async (req) => {
  try {
    // Verificar token secreto en query param para autenticar el webhook
    const url = new URL(req.url);
    const token = url.searchParams.get("token");
    if (token !== Deno.env.get("FRESHDESK_WEBHOOK_SECRET")) {
      return new Response("Unauthorized", { status: 401 });
    }

    const payload = await req.json();

    // Freshdesk envía el ticket con subject y status
    // Solo procesar tickets resueltos o cerrados
    const status = payload?.ticket?.status;
    const subject = payload?.ticket?.subject ?? "";
    if (status !== "resolved" && status !== "closed" && status !== 4 && status !== 5) {
      return new Response(JSON.stringify({ skip: "no es resolved/closed" }), { status: 200 });
    }

    // Extraer folio del asunto: "[NEXUS] Solicitud SAP-2026-XXXXXX"
    const match = subject.match(/\[NEXUS\]\sSolicitud\s(\S+)/);
    if (!match) {
      return new Response(JSON.stringify({ skip: "no se encontró folio en subject" }), { status: 200 });
    }
    const folio = match[1];

    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
    );

    // Obtener la solicitud
    const { data: sol, error } = await supabase
      .from("solicitudes")
      .select("folio, solicitante_nombre, solicitante_email, planillas, historial, estado")
      .eq("folio", folio)
      .single();

    if (error || !sol) {
      return new Response(JSON.stringify({ error: "Solicitud no encontrada", folio }), { status: 404 });
    }

    // Evitar actualizar si ya está Aplicada
    if (sol.estado === "Aplicada") {
      return new Response(JSON.stringify({ skip: "ya estaba Aplicada" }), { status: 200 });
    }

    // Actualizar estado a Aplicada
    const historial = [...(sol.historial || []), { estado: "Aplicada", fecha: new Date().toISOString() }];
    await supabase
      .from("solicitudes")
      .update({ estado: "Aplicada", historial, updated_at: new Date().toISOString() })
      .eq("folio", folio);

    // Enviar email al solicitante
    const transporter = nodemailer.createTransport({
      host: "smtp.gmail.com",
      port: 587,
      secure: false,
      auth: {
        user: Deno.env.get("GMAIL_USER"),
        pass: Deno.env.get("GMAIL_APP_PASSWORD"),
      },
    });

    const planillasTexto = (sol.planillas || [])
      .map((p: { nombre: string; filas: unknown[] }) => `  • ${p.nombre}: ${p.filas.length} filas`)
      .join("\n");

    await transporter.sendMail({
      from: `"NEXUS · Arcoprime" <${Deno.env.get("GMAIL_USER")}>`,
      to: sol.solicitante_email,
      subject: `[NEXUS] Solicitud ${folio} — Aplicada en SAP ✓`,
      text: `Hola ${sol.solicitante_nombre},

Tu solicitud ${folio} ha sido APLICADA en SAP.

Planillas procesadas:
${planillasTexto}

Los cambios ya están disponibles en el sistema. Puedes verificarlo en el portal NEXUS → Solicitudes.

Saludos,
NEXUS · Arcoprime`,
    });

    return new Response(JSON.stringify({ ok: true, folio }), {
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), { status: 500 });
  }
});
