import { createClient } from '@supabase/supabase-js';

// Replace these placeholders with your actual keys from your Supabase dashboard
const supabaseUrl = 'https://elhqbkkctsebzqonbwld.supabase.co/'; // <--- Put your URL here
const supabaseAnonKey = 'sb_publishable_GSyS72mYZjweS_xd6vlYMQ_59aJr_cn'; // <--- Put your Anon Key here

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
