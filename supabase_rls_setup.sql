-- Activer formellement RLS sur la table
ALTER TABLE public.customer_pricing ENABLE ROW LEVEL SECURITY;

-- 1. Autoriser la **lecture (SELECT)** de la table de prix à TOUT LE MONDE
-- L'Appli (qui est neutre à la base) doit pouvoir lire les prix par dfaut au chargement
CREATE POLICY "Enable read access for all users" 
ON public.customer_pricing 
FOR SELECT 
USING (true);

-- 2. Autoriser l'**insertion (INSERT)** UNIQUEMENT aux utilisateurs connectés (authenticated)
CREATE POLICY "Enable insert for authenticated users only" 
ON public.customer_pricing 
FOR INSERT 
TO authenticated 
WITH CHECK (true);

-- 3. Autoriser la **modification (UPDATE)** UNIQUEMENT aux utilisateurs connectés (authenticated)
CREATE POLICY "Enable update for authenticated users only" 
ON public.customer_pricing 
FOR UPDATE 
TO authenticated 
USING (true)
WITH CHECK (true);

-- (Optionnel) 4. Autoriser la suppression (DELETE) UNIQUEMENT aux utilisateurs connectés
CREATE POLICY "Enable delete for authenticated users only" 
ON public.customer_pricing 
FOR DELETE 
TO authenticated 
USING (true);
