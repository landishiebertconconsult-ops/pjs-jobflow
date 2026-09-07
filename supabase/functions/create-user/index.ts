import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import { createClient } from "@supabase/supabase-js";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json",
    },
  });
}

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const { email, password, full_name, role } = await req.json();

    if (!email || !password || !full_name || !role) {
      return jsonResponse(
        {
          success: false,
          error: "Name, email, password, and role are required.",
        },
        400
      );
    }

    const supabaseUrl = Deno.env.get("PROJECT_URL");
    const serviceRoleKey = Deno.env.get("SERVICE_ROLE_KEY");

    if (!supabaseUrl || !serviceRoleKey) {
      return jsonResponse(
        {
          success: false,
          error: "Server configuration is missing.",
        },
        500
      );
    }

    const adminClient = createClient(supabaseUrl, serviceRoleKey);

    const cleanEmail = String(email).trim().toLowerCase();
    const cleanName = String(full_name).trim();

    // JobFlow UI -> Supabase database role
    const roleMap: Record<string, string> = {
      Admin: "Admin",
      "Project Manager": "Project Manager",
      Foreman: "Foreman",
      Crew: "Crew Member",
      "Crew Member": "Crew Member",

      admin: "Admin",
      project_manager: "Project Manager",
      foreman: "Foreman",
      crew: "Crew Member",
      crew_member: "Crew Member",
    };

    const roleText = String(role);
    const cleanRole =
      roleMap[roleText] ||
      roleMap[roleText.toLowerCase()] ||
      "Crew Member";

    // Create the Authentication account
    const { data: authData, error: authError } =
      await adminClient.auth.admin.createUser({
        email: cleanEmail,
        password: String(password),
        email_confirm: true,
        user_metadata: {
          full_name: cleanName,
          role: cleanRole,
        },
      });

    if (authError) {
      return jsonResponse(
        {
          success: false,
          error: authError.message,
          stage: "auth",
        },
        400
      );
    }

    const user = authData.user;

    if (!user) {
      return jsonResponse(
        {
          success: false,
          error: "Supabase did not return the newly created user.",
        },
        500
      );
    }

    // Create/update the matching JobFlow profile
    const { data: profile, error: profileError } = await adminClient
      .from("profiles")
      .upsert(
        {
          id: user.id,
          email: cleanEmail,
          full_name: cleanName,
          role: cleanRole,
          points: 0,
          active: true,
        },
        {
          onConflict: "id",
        }
      )
      .select()
      .single();

    if (profileError) {
      // Don't leave an orphan Authentication user if profile creation fails
      await adminClient.auth.admin.deleteUser(user.id);

      return jsonResponse(
        {
          success: false,
          error: profileError.message,
          stage: "profile",
        },
        400
      );
    }

    return jsonResponse({
      success: true,
      user: {
        id: user.id,
        email: user.email,
      },
      profile,
    });
  } catch (err) {
    return jsonResponse(
      {
        success: false,
        error: err instanceof Error ? err.message : "Unknown server error.",
      },
      500
    );
  }
});