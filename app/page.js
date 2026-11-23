"use client";
import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  ShoppingCart, Package, History, Upload, Download, 
  Plus, Minus, Trash2, DollarSign, CreditCard, ArrowRightLeft, Search, CheckCircle2 
} from 'lucide-react';

export default function POSApp() {
  const [activeTab, setActiveTab] = useState('pos');
  const [inventory, setInventory] = useState([]);
  const [cart, setCart] = useState([]);
  const [sales, setSales] = useState([]);
  const [search, setSearch] = useState('');
  
  // Estados de pago
  const [paymentMethod, setPaymentMethod] = useState('Efectivo');
  const [cashGiven, setCashGiven] = useState('');
  const [saleSuccess, setSaleSuccess] = useState(false);

  // Estados para agregar producto manual
  const [newProdName, setNewProdName] = useState('');
  const [newProdPrice, setNewProdPrice] = useState('');
  const [newProdStock, setNewProdStock] = useState('');

  // Cargar datos guardados al iniciar
  useEffect(() => {
    const savedInv = localStorage.getItem('chele_inventory');
    const savedSales = localStorage.getItem('chele_sales');
    if (savedInv) {
      setInventory(JSON.parse(savedInv));
    } else {
      // Datos de ejemplo iniciales
      setInventory([
        { id: 1, name: 'Minuta Tradicional', price: 1.50, stock: 50 },
        { id: 2, name: 'Minuta Especial con Leche', price: 2.50, stock: 30 },
        { id: 3, name: 'Minuta de Tamarindo', price: 1.75, stock: 40 },
      ]);
    }
    if (savedSales) setSales(JSON.parse(savedSales));
  }, []);

  // Guardar automáticamente en el navegador
  useEffect(() => {
    localStorage.setItem('chele_inventory', JSON.stringify(inventory));
  }, [inventory]);

  useEffect(() => {
    localStorage.setItem('chele_sales', JSON.stringify(sales));
  }, [sales]);

  // --- LÓGICA DEL CARRITO Y VENTA ---
  const addToCart = (product) => {
    if (product.stock <= 0) return;
    const existing = cart.find((item) => item.id === product.id);
    if (existing) {
      if (existing.qty < product.stock) {
        setCart(cart.map((item) => item.id === product.id ? { ...item, qty: item.qty + 1 } : item));
      }
    } else {
      setCart([...cart, { ...product, qty: 1 }]);
    }
  };

  const updateCartQty = (id, delta) => {
    const prod = inventory.find((p) => p.id === id);
    setCart(cart.map((item) => {
      if (item.id === id) {
        const newQty = item.qty + delta;
        if (newQty <= 0) return null;
        if (prod && newQty > prod.stock) return item;
        return { ...item, qty: newQty };
      }
      return item;
    }).filter(Boolean));
  };

  const cartTotal = cart.reduce((sum, item) => sum + item.price * item.qty, 0);
  const changeAmount = paymentMethod === 'Efectivo' && cashGiven ? parseFloat(cashGiven) - cartTotal : 0;

  const handleCompleteSale = () => {
    if (cart.length === 0) return;
    if (paymentMethod === 'Efectivo' && (parseFloat(cashGiven) < cartTotal || !cashGiven)) {
      alert('El monto recibido en efectivo debe ser igual o mayor al total.');
      return;
    }

    // 1. Descontar inventario
    const updatedInventory = inventory.map((invItem) => {
      const soldItem = cart.find((c) => c.id === invItem.id);
      if (soldItem) {
        return { ...invItem, stock: Math.max(0, invItem.stock - soldItem.qty) };
      }
      return invItem;
    });
    setInventory(updatedInventory);

    // 2. Registrar venta
    const newSale = {
      id: Date.now(),
      date: new Date().toLocaleString(),
      items: cart,
      total: cartTotal,
      paymentMethod,
      cashGiven: paymentMethod === 'Efectivo' ? parseFloat(cashGiven) : cartTotal,
      change: paymentMethod === 'Efectivo' ? changeAmount : 0,
    };
    setSales([newSale, ...sales]);

    // 3. Limpiar caja
    setCart([]);
    setCashGiven('');
    setSaleSuccess(true);
    setTimeout(() => setSaleSuccess(false), 3000);
  };

  // --- LÓGICA DE EXCEL E INVENTARIO ---
  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);

      const formatted = data.map((row, index) => ({
        id: Date.now() + index,
        name: row.Producto || row.producto || row.Nombre || row.nombre || `Producto ${index + 1}`,
        price: parseFloat(row.Precio || row.precio || 0),
        stock: parseInt(row.Stock || row.stock || row.Cantidad || row.cantidad || 0, 10),
      }));

      setInventory(formatted);
      alert(`¡Se cargaron ${formatted.length} productos desde Excel!`);
    };
    reader.readAsBinaryString(file);
  };

  const exportInventoryExcel = () => {
    const dataToExport = inventory.map(p => ({
      Producto: p.name,
      Precio: p.price,
      Stock: p.stock
    }));
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inventario");
    XLSX.writeFile(wb, "Inventario_Actualizado.xlsx");
  };

  const exportSalesExcel = () => {
    const dataToExport = sales.map(s => ({
      Fecha: s.date,
      Productos: s.items.map(i => `${i.qty}x ${i.name}`).join(', '),
      Total: s.total.toFixed(2),
      Metodo_Pago: s.paymentMethod,
      Efectivo_Recibido: s.cashGiven.toFixed(2),
      Cambio: s.change.toFixed(2)
    }));
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ventas");
    XLSX.writeFile(wb, "Reporte_Ventas.xlsx");
  };

  const handleAddManualProduct = (e) => {
    e.preventDefault();
    if (!newProdName || !newProdPrice || !newProdStock) return;
    setInventory([
      ...inventory,
      {
        id: Date.now(),
        name: newProdName,
        price: parseFloat(newProdPrice),
        stock: parseInt(newProdStock, 10)
      }
    ]);
    setNewProdName('');
    setNewProdPrice('');
    setNewProdStock('');
  };

  const filteredInventory = inventory.filter(p => 
    p.name.toLowerCase().includes(search.toLowerCase())
  );

  const totalRevenue = sales.reduce((acc, s) => acc + s.total, 0);

  return (
    <div className="min-h-screen flex flex-col">
      {/* Barra Superior */}
      <header className="bg-blue-700 text-white shadow-md">
        <div className="max-w-7xl mx-auto px-4 py-4 flex flex-wrap justify-between items-center gap-4">
          <h1 className="text-2xl font-bold tracking-tight">🍧 Minutas el Chele — POS</h1>
          <nav className="flex gap-2">
            <button 
              onClick={() => setActiveTab('pos')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg font-medium transition ${activeTab === 'pos' ? 'bg-white text-blue-700 shadow' : 'bg-blue-800 hover:bg-blue-600'}`}
            >
              <ShoppingCart size={18} /> Punto de Venta
            </button>
            <button 
              onClick={() => setActiveTab('inventory')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg font-medium transition ${activeTab === 'inventory' ? 'bg-white text-blue-700 shadow' : 'bg-blue-800 hover:bg-blue-600'}`}
            >
              <Package size={18} /> Inventario
            </button>
            <button 
              onClick={() => setActiveTab('sales')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg font-medium transition ${activeTab === 'sales' ? 'bg-white text-blue-700 shadow' : 'bg-blue-800 hover:bg-blue-600'}`}
            >
              <History size={18} /> Ventas ({sales.length})
            </button>
          </nav>
        </div>
      </header>

      {/* Contenido Principal */}
      <main className="max-w-7xl w-full mx-auto p-4 flex-1">
        
        {/* PESTAÑA 1: PUNTO DE VENTA */}
        {activeTab === 'pos' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Catálogo de Productos */}
            <div className="lg:col-span-2 space-y-4">
              <div className="relative">
                <Search className="absolute left-3 top-3 text-slate-400" size={20} />
                <input
                  type="text"
                  placeholder="Buscar producto..."
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  className="w-full pl-10 pr-4 py-2.5 rounded-xl border border-slate-300 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
                {filteredInventory.map((product) => (
                  <button
                    key={product.id}
                    onClick={() => addToCart(product)}
                    disabled={product.stock <= 0}
                    className={`p-4 rounded-xl border text-left transition flex flex-col justify-between h-32 ${
                      product.stock > 0 
                        ? 'bg-white border-slate-200 hover:border-blue-500 hover:shadow-md active:scale-95' 
                        : 'bg-slate-200 border-slate-300 opacity-60 cursor-not-allowed'
                    }`}
                  >
                    <span className="font-semibold text-slate-800 line-clamp-2">{product.name}</span>
                    <div className="flex justify-between items-end w-full mt-2">
                      <span className="text-lg font-bold text-blue-600">${product.price.toFixed(2)}</span>
                      <span className={`text-xs px-2 py-1 rounded-full font-medium ${
                        product.stock > 5 ? 'bg-green-100 text-green-700' : product.stock > 0 ? 'bg-amber-100 text-amber-700' : 'bg-red-100 text-red-700'
                      }`}>
                        {product.stock > 0 ? `Stock: ${product.stock}` : 'Agotado'}
                      </span>
                    </div>
                  </button>
                ))}
              </div>
            </div>

            {/* Panel de Cobro */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-col justify-between h-fit sticky top-4">
              <div>
                <h2 className="text-lg font-bold border-b pb-3 mb-3 flex items-center justify-between">
                  <span>Orden Actual</span>
                  {cart.length > 0 && (
                    <button onClick={() => setCart([])} className="text-xs text-red-500 hover:underline">Limpiar</button>
                  )}
                </h2>

                {cart.length === 0 ? (
                  <p className="text-slate-400 text-center py-8">Selecciona productos del catálogo para vender</p>
                ) : (
                  <div className="space-y-3 max-h-64 overflow-y-auto pr-1 mb-4">
                    {cart.map((item) => (
                      <div key={item.id} className="flex justify-between items-center bg-slate-50 p-2.5 rounded-lg">
                        <div className="flex-1">
                          <p className="font-medium text-sm">{item.name}</p>
                          <p className="text-xs text-slate-500">${item.price.toFixed(2)} c/u</p>
                        </div>
                        <div className="flex items-center gap-2">
                          <button onClick={() => updateCartQty(item.id, -1)} className="p-1 bg-slate-200 rounded hover:bg-slate-300"><Minus size={14} /></button>
                          <span className="font-bold w-6 text-center text-sm">{item.qty}</span>
                          <button onClick={() => updateCartQty(item.id, 1)} className="p-1 bg-slate-200 rounded hover:bg-slate-300"><Plus size={14} /></button>
                        </div>
                        <div className="w-16 text-right font-semibold text-sm">
                          ${(item.price * item.qty).toFixed(2)}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div className="border-t pt-4 space-y-4">
                <div className="flex justify-between items-center text-xl font-bold">
                  <span>Total a Pagar:</span>
                  <span className="text-blue-700">${cartTotal.toFixed(2)}</span>
                </div>

                {/* Método de Pago */}
                <div>
                  <label className="block text-xs font-semibold uppercase text-slate-500 mb-2">Método de Pago</label>
                  <div className="grid grid-cols-3 gap-2">
                    {[
                      { id: 'Efectivo', icon: DollarSign },
                      { id: 'Transferencia', icon: ArrowRightLeft },
                      { id: 'Tarjeta', icon: CreditCard }
                    ].map(({ id, icon: Icon }) => (
                      <button
                        key={id}
                        type="button"
                        onClick={() => setPaymentMethod(id)}
                        className={`flex flex-col items-center justify-center py-2 px-1 rounded-lg border text-xs font-medium transition ${
                          paymentMethod === id 
                            ? 'bg-blue-50 border-blue-600 text-blue-700' 
                            : 'border-slate-200 text-slate-600 hover:bg-slate-50'
                        }`}
                      >
                        <Icon size={16} className="mb-1" />
                        {id}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Calculadora de Efectivo */}
                {paymentMethod === 'Efectivo' && (
                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 space-y-2">
                    <div className="flex justify-between items-center">
                      <label className="text-sm font-medium text-slate-700">Efectivo recibido ($):</label>
                      <input
                        type="number"
                        step="0.01"
                        placeholder="0.00"
                        value={cashGiven}
                        onChange={(e) => setCashGiven(e.target.value)}
                        className="w-28 text-right font-bold p-1.5 rounded border border-slate-300 bg-white"
                      />
                    </div>
                    <div className="flex justify-between items-center pt-1 border-t border-slate-200">
                      <span className="text-sm font-medium text-slate-600">Cambio a devolver:</span>
                      <span className={`text-lg font-bold ${changeAmount < 0 ? 'text-red-500' : 'text-green-600'}`}>
                        ${changeAmount > 0 ? changeAmount.toFixed(2) : '0.00'}
                      </span>
                    </div>
                  </div>
                )}

                <button
                  onClick={handleCompleteSale}
                  disabled={cart.length === 0}
                  className="w-full py-3 bg-green-600 hover:bg-green-700 disabled:bg-slate-300 text-white font-bold rounded-xl shadow transition flex items-center justify-center gap-2"
                >
                  Guardar Venta
                </button>

                {saleSuccess && (
                  <div className="flex items-center justify-center gap-2 text-green-700 bg-green-100 p-2.5 rounded-lg text-sm font-medium">
                    <CheckCircle2 size={18} /> ¡Venta registrada e inventario actualizado!
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* PESTAÑA 2: INVENTARIO Y EXCEL */}
        {activeTab === 'inventory' && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div>
                <h2 className="text-lg font-bold">Gestión de Inventario</h2>
                <p className="text-sm text-slate-500">Sube un Excel con las columnas: <strong>Producto</strong>, <strong>Precio</strong> y <strong>Stock</strong>.</p>
              </div>
              <div className="flex flex-wrap gap-3">
                <label className="cursor-pointer flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl font-medium transition">
                  <Upload size={18} /> Cargar Excel (.xlsx)
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={handleExcelUpload} className="hidden" />
                </label>
                <button onClick={exportInventoryExcel} className="flex items-center gap-2 bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-xl font-medium transition">
                  <Download size={18} /> Descargar Inventario
                </button>
              </div>
            </div>

            {/* Formulario agregar producto manual */}
            <form onSubmit={handleAddManualProduct} className="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap gap-3 items-end">
              <div className="flex-1 min-w-[200px]">
                <label className="block text-xs font-semibold text-slate-500 mb-1">Nombre del Producto</label>
                <input type="text" required value={newProdName} onChange={e => setNewProdName(e.target.value)} placeholder="Ej. Minuta de Fresa" className="w-full p-2 border rounded-lg" />
              </div>
              <div className="w-32">
                <label className="block text-xs font-semibold text-slate-500 mb-1">Precio ($)</label>
                <input type="number" step="0.01" required value={newProdPrice} onChange={e => setNewProdPrice(e.target.value)} placeholder="1.50" className="w-full p-2 border rounded-lg" />
              </div>
              <div className="w-32">
                <label className="block text-xs font-semibold text-slate-500 mb-1">Stock Inicial</label>
                <input type="number" required value={newProdStock} onChange={e => setNewProdStock(e.target.value)} placeholder="20" className="w-full p-2 border rounded-lg" />
              </div>
              <button type="submit" className="bg-green-600 hover:bg-green-700 text-white px-5 py-2 rounded-lg font-medium flex items-center gap-1">
                <Plus size={18} /> Agregar
              </button>
            </form>

            {/* Tabla de Inventario */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
              <table className="w-full text-left border-collapse">
                <thead className="bg-slate-50 border-b">
                  <tr>
                    <th className="p-4 font-semibold text-sm text-slate-600">Producto</th>
                    <th className="p-4 font-semibold text-sm text-slate-600">Precio ($)</th>
                    <th className="p-4 font-semibold text-sm text-slate-600">Stock Disponible</th>
                    <th className="p-4 font-semibold text-sm text-slate-600 text-right">Acciones</th>
                  </tr>
                </thead>
                <tbody className="divide-y">
                  {inventory.map((item) => (
                    <tr key={item.id} className="hover:bg-slate-50">
                      <td className="p-4 font-medium">{item.name}</td>
                      <td className="p-4">
                        <input
                          type="number"
                          step="0.01"
                          value={item.price}
                          onChange={(e) => setInventory(inventory.map(p => p.id === item.id ? { ...p, price: parseFloat(e.target.value) || 0 } : p))}
                          className="w-24 p-1 border rounded"
                        />
                      </td>
                      <td className="p-4">
                        <input
                          type="number"
                          value={item.stock}
                          onChange={(e) => setInventory(inventory.map(p => p.id === item.id ? { ...p, stock: parseInt(e.target.value, 10) || 0 } : p))}
                          className="w-20 p-1 border rounded font-bold"
                        />
                      </td>
                      <td className="p-4 text-right">
                        <button
                          onClick={() => setInventory(inventory.filter(p => p.id !== item.id))}
                          className="text-red-500 hover:bg-red-50 p-2 rounded-lg"
                        >
                          <Trash2 size={18} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* PESTAÑA 3: HISTORIAL DE VENTAS */}
        {activeTab === 'sales' && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div>
                <h2 className="text-lg font-bold">Historial de Ventas</h2>
                <p className="text-2xl font-bold text-green-600 mt-1">Total Acumulado: ${totalRevenue.toFixed(2)}</p>
              </div>
              <button onClick={exportSalesExcel} className="flex items-center gap-2 bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-xl font-medium transition">
                <Download size={18} /> Exportar Ventas a Excel
              </button>
            </div>

            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
              {sales.length === 0 ? (
                <p className="text-center text-slate-400 py-10">Aún no se han registrado ventas.</p>
              ) : (
                <table className="w-full text-left border-collapse">
                  <thead className="bg-slate-50 border-b">
                    <tr>
                      <th className="p-4 font-semibold text-sm text-slate-600">Fecha y Hora</th>
                      <th className="p-4 font-semibold text-sm text-slate-600">Detalle</th>
                      <th className="p-4 font-semibold text-sm text-slate-600">Método</th>
                      <th className="p-4 font-semibold text-sm text-slate-600">Total</th>
                      <th className="p-4 font-semibold text-sm text-slate-600">Cambio</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {sales.map((sale) => (
                      <tr key={sale.id} className="hover:bg-slate-50">
                        <td className="p-4 text-sm text-slate-500">{sale.date}</td>
                        <td className="p-4 text-sm font-medium">
                          {sale.items.map(i => `${i.qty}x ${i.name}`).join(', ')}
                        </td>
                        <td className="p-4">
                          <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-blue-100 text-blue-700">
                            {sale.paymentMethod}
                          </span>
                        </td>
                        <td className="p-4 font-bold text-slate-800">${sale.total.toFixed(2)}</td>
                        <td className="p-4 text-sm text-slate-500">
                          {sale.paymentMethod === 'Efectivo' ? `$${sale.change.toFixed(2)}` : '-'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}

      </main>
    </div>
  );
}
