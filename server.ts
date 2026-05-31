import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  // API Routes
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  app.post("/api/generate-cards", async (req, res) => {
    const { topic, count } = req.body;
    if (!topic) {
      return res.status(400).json({ error: "Falta el tema (topic)" });
    }
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: "GEMINI_API_KEY no está configurada en el servidor." });
    }
    try {
      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });
      const cardsCount = count || 50;
      const prompt = `Genera exactamente ${cardsCount} tarjetas para un juego de Tabú sobre el tema: "${topic}". 
      Cada tarjeta representa un concepto graduado en nivel de complejidad o logro según la escala de la palabra (es decir, de más fácil a más compleja):
      - "Inicio" (conceptos muy sencillos, términos comunes y fáciles de adivinar).
      - "Proceso" (conceptos intermedios o procesos básicos, un poco más específicos).
      - "Logrado" (palabras técnicas, teóricas o de mayor nivel conceptual).
      - "Destacado" (conceptos avanzados, abstractos, muy retadores o especializados).

      Asegúrate de distribuir las tarjetas equitativamente entre los cuatro niveles de logro en base a su dificultad o complejidad técnica.

      Cada objeto de tarjeta debe tener:
      1. "nivelLogro": El nivel de logro ("Inicio", "Proceso", "Logrado" o "Destacado").
      2. "categoria": La categoría temática (debe ser "${topic}" o algo muy relacionado).
      3. "secreto": La palabra que el jugador debe adivinar.
      4. "tabu": Una lista de 4 palabras prohibidas que no se pueden usar para describir la palabra secreta.
      5. "pista": Una pista breve (máximo 15 palabras) estilo socrático o una definición indirecta para ayudar en español.
      
      Responde estrictamente en formato JSON como un arreglo de objetos.`;

      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: prompt,
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                nivelLogro: { type: Type.STRING },
                secreto: { type: Type.STRING },
                tabu: { 
                  type: Type.ARRAY,
                  items: { type: Type.STRING }
                },
                categoria: { type: Type.STRING },
                pista: { type: Type.STRING }
              },
              required: ["nivelLogro", "secreto", "tabu", "categoria", "pista"]
            }
          }
        }
      });

      if (!response.text) {
        throw new Error("La IA no devolvió ninguna respuesta.");
      }

      const cleanJson = response.text.replace(/```json/g, "").replace(/```/g, "").trim();
      const parsed = JSON.parse(cleanJson);
      res.json(parsed);
    } catch (error: any) {
      console.error("Error generating cards in server:", error);
      res.status(500).json({ error: error.message || "Error al generar tarjetas con IA" });
    }
  });

  app.post("/api/socratic-hint", async (req, res) => {
    const { word } = req.body;
    if (!word) {
      return res.status(400).json({ error: "Falta la palabra (word)" });
    }
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: "GEMINI_API_KEY no está configurada en el servidor." });
    }
    try {
      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });
      const prompt = `Actúa como un profesor Socrático. Dame una pista indirecta, una definición conceptual o un dato de contexto breve sobre la palabra o concepto "${word}" en español para ayudar a alguien a adivinarla sin usar la palabra en sí ni sus conceptos más obvios. Debe ser educativo, indirecto y desafiante. Sé breve. Máximo 20 palabras.`;

      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: prompt
      });

      res.json({ hint: response.text || "" });
    } catch (error: any) {
      console.error("Error generating socratic hint in server:", error);
      res.status(500).json({ error: error.message || "Error al generar pista" });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
