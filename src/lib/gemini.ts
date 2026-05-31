import { CardData } from "../types";

// Llama al servidor de Node/Express para mantener oculta la API key del navegador de forma segura
export async function generateTabooCards(topic: string, count: number = 50): Promise<CardData[]> {
  try {
    const response = await fetch("/api/generate-cards", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ topic, count })
    });

    if (!response.ok) {
      const errBody = await response.json().catch(() => ({}));
      throw new Error(errBody.error || `Error del servidor (${response.status})`);
    }

    const cards = await response.json();
    
    // Asigna IDs únicos a cada elemento generado para soportar el Inline CRUD en caliente
    return cards.map((card: any, idx: number) => ({
      ...card,
      id: `ai-${idx}-${Date.now()}`
    }));
  } catch (error: any) {
    console.error("Error calling generateTabooCards proxy:", error);
    throw new Error(error.message || "No se pudo comunicar con el servidor de generación.");
  }
}

export async function getSocraticHint(word: string): Promise<string> {
  try {
    const response = await fetch("/api/socratic-hint", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ word })
    });

    if (!response.ok) {
      const errBody = await response.json().catch(() => ({}));
      throw new Error(errBody.error || `Error del servidor (${response.status})`);
    }

    const data = await response.json();
    return data.hint || "No se pudo obtener una pista válida.";
  } catch (error: any) {
    console.error("Error calling getSocraticHint proxy:", error);
    return "No se pudo obtener una pista en este momento.";
  }
}
