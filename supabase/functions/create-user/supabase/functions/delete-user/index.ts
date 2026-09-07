import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

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
  // Handle browser preflight request
  if (req.method === "OPTIONS") {
    return new Response("ok", {
      headers: corsHeaders,
    });
  }

  if (req.method !== "POST") {
    return jsonResponse(
      {
        success: false,
        error: "Method not allowed.",
      },
      405
    );
  }

  try {
    // Get Supabase environment variables
    const supabaseUrl =
      Deno.env.get("SUPABASE_URL") ??
      Deno.env.get("PROJECT_URL");

    const serviceRoleKey =
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ??
      Deno.env.get("SERVICE_ROLE_KEY");

    if (!supabaseUrl || !serviceRoleKey) {
      return jsonResponse(
        {
          success: false,
          error: "Supabase server environment variables are missing.",
        },
        500
      );
    }

    // Create admin Supabase client
    const adminClient = createClient(
      supabaseUrl,
      serviceRoleKey,
      {
        auth: {
          autoRefreshToken: false,
          persistSession: false,
        },
      }
    );

    // Verify whoever called this function is logged in
    const authHeader = req.headers.get("Authorization");

    if (!authHeader) {
      return jsonResponse(
        {
          success: false,
          error: "Not authenticated.",
        },
        401
      );
    }

    const token = authHeader.replace("Bearer ", "");

    const {
      data: { user: requestingUser },
      error: requestingUserError,
    } = await adminClient.auth.getUser(token);

    if (requestingUserError || !requestingUser) {
      return jsonResponse(
        {
          success: false,
          error: "Invalid login session.",
        },
        401
      );
    }

    // Check that the person performing the delete is an Admin
    const {
      data: requestingProfile,
      error: requestingProfileError,
    } = await adminClient
      .from("profiles")
      .select("id, role")
      .eq("id", requestingUser.id)
      .single();

    if (requestingProfileError || !requestingProfile) {
      return jsonResponse(
        {
          success: false,
          error: "Could not verify administrator permissions.",
        },
        403
      );
    }

    if (requestingProfile.role !== "Admin") {
      return jsonResponse(
        {
          success: false,
          error: "Only an Admin can delete users.",
        },
        403
      );
    }

    // Read requested user ID
    const body = await req.json();
    const userId = body?.user_id;

    if (!userId || typeof userId !== "string") {
      return jsonResponse(
        {
          success: false,
          error: "Missing user_id.",
        },
        400
      );
    }

    // Prevent an Admin from accidentally deleting their own account
    if (userId === requestingUser.id) {
      return jsonResponse(
        {
          success: false,
          error: "You cannot delete your own account.",
        },
        400
      );
    }

    // Verify target Auth user exists
    const {
      data: targetUserData,
      error: targetUserError,
    } = await adminClient.auth.admin.getUserById(userId);

    if (targetUserError || !targetUserData?.user) {
      return jsonResponse(
        {
          success: false,
          error:
            targetUserError?.message ??
            "User not found in Supabase Authentication.",
        },
        404
      );
    }

    // Delete Auth user.
    // If profiles.id has ON DELETE CASCADE this also removes the profile.
    const { error: deleteAuthError } =
      await adminClient.auth.admin.deleteUser(userId);

    if (deleteAuthError) {
      return jsonResponse(
        {
          success: false,
          error: deleteAuthError.message,
          stage: "delete_auth_user",
        },
        400
      );
    }

    // Also remove profile explicitly in case cascade is not configured.
    const { error: deleteProfileError } = await adminClient
      .from("profiles")
      .delete()
      .eq("id", userId);

    if (deleteProfileError) {
      return jsonResponse(
        {
          success: false,
          error: deleteProfileError.message,
          stage: "delete_profile",
        },
        400
      );
    }

    return jsonResponse({
      success: true,
      deleted_user_id: userId,
    });
  } catch (err) {
    console.error("delete-user error:", err);

    return jsonResponse(
      {
        success: false,
        error:
          err instanceof Error
            ? err.message
            : "Unknown server error.",
      },
      500
    );
  }
});