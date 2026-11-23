export const metadata = {
  title: 'Minutas el Chele - Punto de Venta',
  description: 'Sistema de ventas e inventario',
}

export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <head>
        <script src="https://cdn.tailwindcss.com"></script>
      </head>
      <body className="bg-slate-100 text-slate-800 antialiased">
        {children}
      </body>
    </html>
  )
}
