import "dotenv/config";
import express from "express";
import cors from "cors";
import https from "https";
import fs from "fs";
import path from "path";
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());

const rootPath = path.join(__dirname, '..');
app.use(express.static(rootPath));

// ===============================
// ENDPOINT PRINCIPAL
// ===============================
app.post("/ia", async (req, res) => {
    try {
        const { prompt } = req.body;

        if (!prompt) {
            throw new Error("No se recibió 'prompt' en el cuerpo de la petición");
        }

        const apiKey = process.env.GEMINI_API_KEY;
        if (!apiKey) {
            throw new Error("La variable de entorno GEMINI_API_KEY no está configurada");
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=${apiKey}`;

        let geminiRes;
        try {
            geminiRes = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt }] }]
                })
            });
        } catch (fetchError) {
            const esErrorDeRed =
                fetchError.cause?.code === 'ENOTFOUND' ||
                fetchError.cause?.code === 'ECONNREFUSED' ||
                fetchError.cause?.code === 'ECONNRESET' ||
                fetchError.cause?.code === 'ETIMEDOUT' ||
                fetchError.message?.includes('fetch failed') ||
                fetchError.message?.includes('ENOTFOUND') ||
                fetchError.message?.includes('ECONNREFUSED');

            if (esErrorDeRed) {
                return res.status(503).json({
                    error: "Sin conexión a internet",
                    details: "No hay conexión a internet. Verifica tu red e inténtalo de nuevo."
                });
            }

            throw fetchError;
        }

        // CONTROL DE ERRORES DE LA API DE GEMINI
        if (!geminiRes.ok) {
            const errorText = await geminiRes.text();

            let parsedError = {};
            try {
                parsedError = JSON.parse(errorText);
            } catch (e) {}

            const message =
                parsedError?.error?.message ||
                parsedError?.message ||
                errorText;

            const lowerMsg = message.toLowerCase();

            // ERROR DE AUTENTICACIÓN (API KEY GEMINI)
            if (
                geminiRes.status === 401 ||
                geminiRes.status === 403 ||
                lowerMsg.includes("api key") ||
                lowerMsg.includes("invalid") ||
                lowerMsg.includes("authentication")
            ) {
                return res.status(401).json({
                    error: "Error de autenticación: La clave de la API de Gemini no es válida o no está configurada.",
                    details: message
                });
            }

            // RATE LIMIT
            if (geminiRes.status === 429) {
                return res.status(429).json({
                    error: "Límite de consultas excedido",
                    details: message
                });
            }

            return res.status(geminiRes.status).json({
                error: "Error remoto de Gemini API",
                details: message
            });
        }

        const data = await geminiRes.json();
        const contenido = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";

        return res.json({ respuesta: contenido });

    } catch (error) {
        console.error("Error en /ia:", error);

        if (error.message?.includes("429")) {
            return res.status(429).json({
                error: "Límite de consultas excedido",
                details: "Se ha alcanzado el límite de peticiones permitido por el modelo gratuito. Por favor, espera un minuto antes de intentar otra consulta."
            });
        }

        return res.status(500).json({ error: "Error del servidor: " + error.message });
    }
});

// ===============================
// HTTPS
// ===============================
const port = 3001;

try {
    const options = {
        key: fs.readFileSync('C:\\Users\\Win 11\\.office-addin-dev-certs\\localhost.key'),
        cert: fs.readFileSync('C:\\Users\\Win 11\\.office-addin-dev-certs\\localhost.crt')
    };

    https.createServer(options, app).listen(port, () => {
        console.log("Servidor HTTPS activo en https://localhost:" + port);
        console.log("Modelo configurado: gemini-3.1-flash-lite (Google AI Studio)");
        console.log("Endpoint disponible: POST /ia");
    });

} catch (err) {
    console.error("Error cargando certificados HTTPS:", err);
}