# Herramientas de la lista

## Clientes y condiciones comerciales

1. Completá el Excel de clientes (guardalo **fuera** de esta carpeta, porque tiene las claves).
   Columnas:

   | CUIT | RAZON SOCIAL | CLAVE | DESCUENTO | DTO Automotor | DTO Industria | … |
   |---|---|---|---|---|---|---|
   | 20-12345678-6 | CLIENTE SA | *(vacía)* | 10 | 15 | | |

   - **CLAVE** vacía: se genera sola y queda guardada en el mismo Excel, para pasársela al cliente.
   - **DESCUENTO**: % general.
   - **DTO \<Rubro\>**: opcional, % para un rubro puntual. Escribí el rubro igual que en la lista.

2. Generá los archivos cifrados:

   ```
   cd herramientas && npm install && cd ..
   node herramientas/clientes.js "C:\ruta\clientes.xlsx"
   ```

3. Subí los cambios (`git add -A && git commit -m "Clientes" && git push`). Netlify publica solo.

Para dar de baja a un cliente, borrá su fila y volvé a correr el paso 2. Para cambiarle la clave, borrá la clave y volvé a correrlo.
