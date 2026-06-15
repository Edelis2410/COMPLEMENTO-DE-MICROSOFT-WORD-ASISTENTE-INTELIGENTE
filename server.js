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

// Ruta estática para el frontend (opcional)
const rootPath = path.join(__dirname, '..');
app.use(express.static(rootPath));

// Endpoint principal
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

        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${apiKey}`;

        const geminiRes = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{
                    parts: [{ text: prompt }]
                }]
            })
        });

        // CONTROL DE ERRORES DE LA API DE GEMINI
        if (!geminiRes.ok) {
            const errorText = await geminiRes.text();
            console.error("Error desde Gemini API:", errorText);

            // CONTROL DE LÍMITES EXCEDIDOS (Rate Limit - Error 429)
            if (geminiRes.status === 429) {
                return res.status(429).json({
                    error: "Límite de consultas excedido",
                    details: "Se ha alcanzado el límite de peticiones permitido por el modelo gratuito. Por favor, espera un minuto antes de intentar otra consulta."
                });
            }

            return res.status(geminiRes.status).json({
                error: "Error remoto de Gemini API",
                details: errorText
            });
        }

        const data = await geminiRes.json();
        const contenido = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";

        return res.json({ respuesta: contenido });

    } catch (error) {
        console.error("Error en /ia:", error);
        
        // Salvaguarda en el catch por si el error de rate limit viene por otra vía
        if (error.message && error.message.includes("429")) {
            return res.status(429).json({ 
                error: "Límite de consultas excedido",
                details: "Se ha alcanzado el límite de peticiones permitido por el modelo gratuito. Por favor, espera un minuto antes de intentar otra consulta."
            });
        }

        return res.status(500).json({ error: "Error del servidor: " + error.message });
    }
});

// ===============================
// HTTPS CONFIGURADO CON RUTAS FIJAS
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
    console.error("Error cargando certificados HTTPS directamente:", err);
}