using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Errepar.MetadataManager.Process.Models
{
    public class ItemMetadataManager
    {
        public int Id { get; set; }//

        // Campos de la lista (según la imagen)
        public JsonDocument Cambios { get; set; }
        public int? CantActivosProcesados { get; set; }//
        public int? CantActivosSeleccionados { get; set; }//
        public DateTime? Creado { get; set; }//
        public string EjecutadoPor { get; set; }      // Persona o grupo (nombre)
        public string EstadoProceso { get; set; }     // Elección (Pendiente, En Pausa, ...)
        public DateTime? FechaFinalizado { get; set; }
        public DateTime? FechaPendiente { get; set; }
        public string Link { get; set; }              // URL del hipervínculo
        public DateTime? Modificado { get; set; }
        public string Titulo { get; set; }            // Title//
        public string CreadoPor { get; set; }         // Autor (nombre)
        public string ModificadoPor { get; set; }     // Editor (nombre)
        public string Scope { get; set; }     // Editor (nombre)
        public string Adjuntos{ get; set; }
        public List<string> Activos { get; set; }

        
    }
}