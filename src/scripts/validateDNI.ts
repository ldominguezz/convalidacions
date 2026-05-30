// src/utils/validateDNI.ts

const LETRAS = "TRWAGMYFPDXBNJZSQVHLCKE" as const;

interface ValidationResult {
  valid: boolean;
  error?: string;
}

export function validateDNI(dni: string): ValidationResult {
  const doc = dni.trim().toUpperCase();
  if (!/^[0-9]{8}[A-Z]$/.test(doc))
    return { valid: false, error: "Format incorrecte (8 dígits + lletra)." };
  const esperada = LETRAS[parseInt(doc.slice(0, 8), 10) % 23];
  if (doc[8] !== esperada)
    return {
      valid: false,
      error: `Lletra incorrecta. Ha de ser "${esperada}".`,
    };
  return { valid: true };
}

export function validateNIE(nie: string): ValidationResult {
  const doc = nie.trim().toUpperCase();
  if (!/^[XYZ][0-9]{7}[A-Z]$/.test(doc))
    return {
      valid: false,
      error: "Format incorrecte (X/Y/Z + 7 dígits + lletra).",
    };
  const map: Record<string, string> = { X: "0", Y: "1", Z: "2" };
  const num = parseInt(map[doc[0]] + doc.slice(1, 8), 10);
  const esperada = LETRAS[num % 23];
  if (doc[8] !== esperada)
    return {
      valid: false,
      error: `Lletra incorrecta. Ha de ser "${esperada}".`,
    };
  return { valid: true };
}

export function validateDocument(doc: string): ValidationResult {
  const cleaned = doc.trim().toUpperCase();
  if (/^[XYZ]/.test(cleaned)) return validateNIE(cleaned);
  if (/^[0-9]/.test(cleaned)) return validateDNI(cleaned);
  return {
    valid: false,
    error: "Ha de començar amb un dígit (DNI) o X, Y, Z (NIE).",
  };
}
