import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    // Cliente con anon key para verificar identidad del llamante
    const supabaseAnon = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_ANON_KEY")!,
      { global: { headers: { Authorization: req.headers.get("Authorization")! } } }
    );

    // Verificar que el llamante está autenticado y tiene rol datos_maestros
    const { data: { user }, error: authError } = await supabaseAnon.auth.getUser();
    if (authError || !user) {
      return new Response(JSON.stringify({ error: "No autenticado" }), {
        status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const { data: perfil } = await supabaseAnon
      .from("profiles")
      .select("rol")
      .eq("id", user.id)
      .single();

    if (perfil?.rol !== "datos_maestros") {
      return new Response(JSON.stringify({ error: "Sin permisos" }), {
        status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Leer parámetros del body
    const { email, nombre, rol } = await req.json();
    if (!email || !nombre) {
      return new Response(JSON.stringify({ error: "email y nombre son requeridos" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Cliente admin con service role para crear la invitación
    const supabaseAdmin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
    );

    const redirectTo = Deno.env.get("SITE_URL") ?? "https://portal-maestros-sap.vercel.app";

    const { data, error } = await supabaseAdmin.auth.admin.inviteUserByEmail(email, {
      data: { nombre, rol: rol ?? "solicitante" },
      redirectTo,
    });

    if (error) {
      return new Response(JSON.stringify({ error: error.message }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Guardar el rol en profiles (el trigger lo crea, pero puede no haber corrido aún)
    await supabaseAdmin
      .from("profiles")
      .upsert({ id: data.user.id, email, nombre, rol: rol ?? "solicitante", activo: true });

    return new Response(JSON.stringify({ ok: true }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
