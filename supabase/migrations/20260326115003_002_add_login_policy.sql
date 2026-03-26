/*
  # Add login policy for employees table

  1. Changes
    - Add policy to allow anonymous users to SELECT from employees table for login purposes
    - This is needed because the app uses custom authentication, not Supabase Auth

  2. Security
    - Only allows SELECT, not INSERT/UPDATE/DELETE for anonymous users
*/

CREATE POLICY "Allow anonymous login check"
  ON employees
  FOR SELECT
  TO anon
  USING (true);
