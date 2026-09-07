import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import { createClient } from "npm:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function respond(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json",
    },
  });
}

serve(async (req) => {
  // Browser CORS check
  if (req.method === "OPTIONS") {
    return new Response("ok", {
      status: 200,
      headers: corsHeaders,
    });
  }

  if (req.method !== "POST") {
    return respond(
      {
        success: false,
        error: "POST requests only.",
      },
      405
    );
  }

  try {
    // Supabase automatically provides these in hosted Edge Functions.
    // Your custom names are included as fallbacks.
    const supabaseUrl =
      Deno.env.get("SUPABASE_URL") ||
      Deno.env.get("PROJECT_URL");

    const serviceRoleKey =
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ||
      Deno.env.get("SERVICE_ROLE_KEY");

    if (!supabaseUrl) {
      return respond(
        {
          success: false,
          error: "Supabase URL is missing.",
        },
        500
      );
    }

    if (!serviceRoleKey) {
      return respond(
        {
          success: false,
          error: "Supabase service role key is missing.",
        },
        500
      );
    }

    const admin = createClient(
      supabaseUrl,
      serviceRoleKey,
      {
        auth: {
          autoRefreshToken: false,
          persistSession: false,
        },
      }
    );

    // -----------------------------
    // Verify logged-in caller
    // -----------------------------

    const authorization = req.headers.get("Authorization");

    if (!authorization) {
      return respond(
        {
          success: false,
          error: "You must be logged in.",
        },
        401
      );
    }

    const token = authorization.replace(/^Bearer\s+/i, "");

    const {
      data: callerData,
      error: callerError,
    } = await admin.auth.getUser(token);

    if (callerError || !callerData?.user) {
      return respond(
        {
          success: false,
          error:
            callerError?.message ||
            "Unable to verify logged-in user.",
        },
        401
      );
    }

    const caller = callerData.user;

    // -----------------------------
    // Verify caller is Admin
    // -----------------------------

    const {
      data: callerProfile,
      error: callerProfileError,
    } = await admin
      .from("profiles")
      .select("id, role")
      .eq("id", caller.id)
      .single();

    if (callerProfileError || !callerProfile) {
      return respond(
        {
          success: false,
          error:
            callerProfileError?.message ||
            "Administrator profile could not be found.",
        },
        403
      );
    }

    if (callerProfile.role !== "Admin") {
      return respond(
        {
          success: false,
          error: "Only an Admin can delete employee accounts.",
        },
        403
      );
    }

    // -----------------------------
    // Get employee being deleted
    // -----------------------------

    let body: any;

    try {
      body = await req.json();
    } catch {
      return respond(
        {
          success: false,
          error: "Invalid request body.",
        },
        400
      );
    }

    const userId = String(body?.user_id || "").trim();

    if (!userId) {
      return respond(
        {
          success: false,
          error: "Missing user_id.",
        },
        400
      );
    }

    // Prevent deleting yourself
    if (userId === caller.id) {
      return respond(
        {
          success: false,
          error: "You cannot delete your own Admin account.",
        },
        400
      );
    }

    // -----------------------------
    // Confirm target Auth user
    // -----------------------------

    const {
      data: targetData,
      error: targetError,
    } = await admin.auth.admin.getUserById(userId);

    if (targetError) {
      return respond(
        {
          success: false,
          error: targetError.message,
          stage: "find_user",
        },
        400
      );
    }

    if (!targetData?.user) {
      return respond(
        {
          success: false,
          error: "Employee could not be found in Authentication.",
        },
        404
      );
    }

    const deletedEmail = targetData.user.email || null;

    // -----------------------------
    // Delete Auth account
    // -----------------------------

    const { error: authDeleteError } =
      await admin.auth.admin.deleteUser(userId);

    if (authDeleteError) {
      return respond(
        {
          success: false,
          error: authDeleteError.message,
          stage: "delete_auth",
        },
        400
      );
    }

    // -----------------------------
    // Delete matching profile
    // -----------------------------

    const { error: profileDeleteError } = await admin
      .from("profiles")
      .delete()
      .eq("id", userId);

    if (profileDeleteError) {
      return respond(
        {
          success: false,
          error:
            "Authentication account was deleted, but profile cleanup failed: " +
            profileDeleteError.message,
          stage: "delete_profile",
        },
        500
      );
    }

    return respond({
      success: true,
      deleted_user_id: userId,
      deleted_email: deletedEmail,
    });
  } catch (error) {
    console.error("delete-user failed:", error);

    return respond(
      {
        success: false,
        error:
          error instanceof Error
            ? error.message
            : "Unknown delete-user error.",
      },
      500
    );
  }
});