## Funcionamiento Técnico

Clase diseñada para la **persistencia de metadatos** en el modelo de objetos de **CATIA V5 (COM API)**. El núcleo utiliza un algoritmo de **recorrido recursivo (DFS)** sobre la colección `Product.Products`. 

Implementa una lógica de filtrado basada en la propiedad `FullName` del documento para distinguir entre:
* **Componentes:** Instancias internas al archivo padre.
* **Documentos Externos:** Archivos `.CATPart` o `.CATProduct`. 

El uso de `HashSet(Of String)` garantiza la **unicidad en la modificación de archivos**, evitando colisiones de escritura en sesiones con múltiples instancias de la misma referencia. Incluye un manejador de excepciones específico para la interfaz `ReferenceProduct` que previene errores de ejecución ante **vínculos rotos (Broken Links)**.
