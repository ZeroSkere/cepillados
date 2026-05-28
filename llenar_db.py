import random
from datetime import datetime, timedelta
from app import app, db, Sabor, Area, Pedido, MetodoPago, generar_numero_factura

# Datos simulados para llenar la base de datos
NOMBRES = [
    "Carlos Rodríguez", "Lesgleiny Martínez", "Mayerlin González", 
    "Emiro Pérez", "Cordero José", "Zoraida Hernández", 
    "Raquel Díaz", "Abigail Torres", "Juan Morales", 
    "María Fernández", "Pedro Ramírez", "Luis Castillo",
    "Ana Suárez", "Jorge Mendoza", "Carmen Ortega", "Rafael Vargas",
    "Sofía Herrera", "Daniel Rojas", "Laura Medina", "Andrés López"
]

AREAS_DATA = [
    "ALMACEN", "SISTEMAS", "COSMETICOS", "FARMACIA", 
    "RECURSOS HUMANOS", "CAJA", "ADMINISTRACIÓN", "VENTAS",
    "CONTABILIDAD", "GERENCIA"
]

# Sabores con stock, precio y referencia de imagen
SABORES_DATA = [
    {"nombre": "fresa", "stock": 50, "precio": 1.0, "color": "#FF6B6B", "emoji": "🍓"},
    {"nombre": "colita", "stock": 60, "precio": 1.0, "color": "#FF3838", "emoji": "🥤"},
    {"nombre": "crema reina", "stock": 40, "precio": 1.5, "color": "#FFD93D", "emoji": "👑"},
    {"nombre": "limon", "stock": 30, "precio": 1.0, "color": "#6BCB77", "emoji": "🍋"},
    {"nombre": "mantecado", "stock": 45, "precio": 1.5, "color": "#F5E6CC", "emoji": "🍦"},
    {"nombre": "uva", "stock": 35, "precio": 2.0, "color": "#9B59B6", "emoji": "🍇"},
    {"nombre": "chocolate", "stock": 25, "precio": 2.0, "color": "#8B4513", "emoji": "🍫"},
    {"nombre": "coco", "stock": 20, "precio": 1.5, "color": "#DEB887", "emoji": "🥥"},
]

# Métodos de pago predefinidos
METODOS_PAGO_DATA = [
    {
        'nombre': 'Pago Móvil',
        'codigo': 'pago_movil',
        'requiere_capture': True,
        'requiere_monto_bs': True,
        'activo': True
    },
    {
        'nombre': 'Efectivo Bolívares',
        'codigo': 'efectivo_bs',
        'requiere_capture': False,
        'requiere_monto_bs': True,
        'activo': True
    },
    {
        'nombre': 'Efectivo Dólares',
        'codigo': 'efectivo_usd',
        'requiere_capture': False,
        'requiere_monto_bs': False,
        'activo': True
    },
    {
        'nombre': 'Transferencia Bancaria',
        'codigo': 'transferencia',
        'requiere_capture': True,
        'requiere_monto_bs': True,
        'activo': True
    },
    {
        'nombre': 'Zelle',
        'codigo': 'zelle',
        'requiere_capture': True,
        'requiere_monto_bs': False,
        'activo': True
    },
    {
        'nombre': 'Punto de Venta',
        'codigo': 'punto_venta',
        'requiere_capture': False,
        'requiere_monto_bs': True,
        'activo': False  # Inactivo para probar la funcionalidad
    }
]

# Tasa BCV simulada
TASA_BCV_SIMULADA = 36.50

def generar_monto_bs(monto_usd, metodo_codigo):
    """Genera un monto en Bs simulado según el método de pago"""
    if metodo_codigo in ['efectivo_bs', 'pago_movil', 'transferencia', 'punto_venta']:
        # 30% de probabilidad de incluir propina
        if random.random() < 0.3:
            propina = random.uniform(0.5, 3.0)
            return round((monto_usd * TASA_BCV_SIMULADA) + propina, 2)
        else:
            return round(monto_usd * TASA_BCV_SIMULADA, 2)
    return None

def llenar_base_datos():
    with app.app_context():
        print("=" * 70)
        print("🍧 SEMILLERO DE BASE DE DATOS - CEPILLADOS v1.1")
        print("=" * 70)
        
        print("\n🗑️  Limpiando base de datos anterior...")
        db.drop_all()
        db.create_all()
        print("   ✅ Base de datos limpia y recreada")

        # ==================== CREAR ÁREAS ====================
        print("\n🏢 Creando Áreas Organizacionales...")
        areas_creadas = []
        for nombre_area in AREAS_DATA:
            nueva_area = Area(nombre=nombre_area)
            db.session.add(nueva_area)
            areas_creadas.append(nueva_area)
            print(f"   ✅ {nombre_area}")
        db.session.commit()

        # ==================== CREAR MÉTODOS DE PAGO ====================
        print("\n💳 Creando Métodos de Pago...")
        metodos_creados = []
        for metodo_data in METODOS_PAGO_DATA:
            nuevo_metodo = MetodoPago(**metodo_data)
            db.session.add(nuevo_metodo)
            metodos_creados.append(nuevo_metodo)
            estado = "ACTIVO" if metodo_data['activo'] else "INACTIVO"
            iconos = []
            if metodo_data['requiere_capture']:
                iconos.append("📸")
            if metodo_data['requiere_monto_bs']:
                iconos.append("💰")
            print(f"   ✅ {metodo_data['nombre']:25} | {metodo_data['codigo']:15} | {estado} {' '.join(iconos)}")
        db.session.commit()

        # ==================== CREAR SABORES ====================
        print("\n🍧 Creando Sabores con Precios y Stock...")
        sabores_creados = []
        for sabor_data in SABORES_DATA:
            nuevo_sabor = Sabor(
                nombre=sabor_data['nombre'],
                stock_inicial=sabor_data['stock'],
                stock_disponible=sabor_data['stock'],
                precio_usd=sabor_data['precio'],
                imagen=None  # Sin imagen por defecto
            )
            db.session.add(nuevo_sabor)
            sabores_creados.append(nuevo_sabor)
            print(f"   ✅ {sabor_data['emoji']} {sabor_data['nombre'].capitalize():15} | Stock: {sabor_data['stock']:3d} | Precio: ${sabor_data['precio']:.2f}")
        db.session.commit()

        # ==================== CREAR PEDIDOS CON FACTURAS ====================
        print("\n📝 Generando 60 pedidos con números de factura secuenciales...")
        pedidos_creados = 0
        fecha_base = datetime.now() - timedelta(days=7)
        
        # Estados posibles con pesos (probabilidades)
        estados_pago = ["si", "si", "si", "no"]  # 75% pagado, 25% pendiente
        estados_entrega = ["si", "si", "no"]     # 66% entregado, 33% pendiente
        
        # Contador de facturas generadas
        facturas_generadas = []
        
        for i in range(60):
            nombre = random.choice(NOMBRES)
            area = random.choice(areas_creadas)
            sabor = random.choice(sabores_creados)
            cantidad = random.choices([1, 2, 3, 4], weights=[40, 35, 20, 5])[0]
            pago = random.choice(estados_pago)
            entrega = random.choice(estados_entrega)
            
            # Si está pagado, asignar método de pago
            metodo_pago_id = None
            metodo_pago_codigo = 'ninguno'
            monto_pagado_bs = None
            capture_img = None
            
            if pago == "si":
                # Elegir solo métodos activos
                metodos_activos = [m for m in metodos_creados if m.activo]
                if metodos_activos:
                    metodo_elegido = random.choice(metodos_activos)
                    metodo_pago_id = metodo_elegido.id
                    metodo_pago_codigo = metodo_elegido.codigo
                    
                    # Calcular monto en USD
                    monto_usd = cantidad * sabor.precio_usd
                    
                    # Generar monto en Bs si el método lo requiere
                    if metodo_elegido.requiere_monto_bs:
                        monto_pagado_bs = generar_monto_bs(monto_usd, metodo_elegido.codigo)
                    
                    # Simular capture (80% de probabilidad si el método lo requiere)
                    if metodo_elegido.requiere_capture and random.random() < 0.8:
                        capture_img = f"capture_simulado_{i+1}.jpg"

            # Generar fecha aleatoria en los últimos 7 días
            dias_atras = random.randint(0, 7)
            horas_atras = random.randint(0, 23)
            minutos_atras = random.randint(0, 59)
            fecha_pedido = fecha_base + timedelta(
                days=dias_atras, 
                hours=horas_atras, 
                minutes=minutos_atras
            )

            # Descontar del stock si hay disponibilidad
            if sabor.stock_disponible >= cantidad:
                sabor.stock_disponible -= cantidad
                
                # Generar número de factura usando la función del sistema
                numero_factura = generar_numero_factura()
                facturas_generadas.append(numero_factura)
                
                nuevo_pedido = Pedido(
                    numero_factura=numero_factura,  # ✅ Número de factura secuencial
                    nombre=nombre,
                    area_id=area.id,
                    sabor_id=sabor.id,
                    cantidad=cantidad,
                    pago=pago,
                    metodo_pago_id=metodo_pago_id,
                    metodo_pago_codigo=metodo_pago_codigo,
                    monto_pagado_bs=monto_pagado_bs,
                    entrega=entrega,
                    capture_img=capture_img,
                    fecha_registro=fecha_pedido
                )
                db.session.add(nuevo_pedido)
                pedidos_creados += 1
                
                # Mostrar progreso con número de factura
                estado_pago_str = "✅ PAGADO" if pago == "si" else "❌ PENDIENTE"
                metodo_str = metodo_pago_codigo.replace('_', ' ').title() if pago == "si" else "N/A"
                print(f"   #{pedidos_creados:2d} | {numero_factura:15s} | {nombre:20s} | "
                      f"{sabor.nombre.capitalize():12s} x{cantidad} | "
                      f"${cantidad * sabor.precio_usd:.2f} | {estado_pago_str} | "
                      f"{fecha_pedido.strftime('%d/%m %H:%M')}")

        db.session.commit()

        # ==================== RESUMEN FINAL ====================
        print("\n" + "=" * 70)
        print("📊 RESUMEN DE LA BASE DE DATOS GENERADA")
        print("=" * 70)
        print(f"   🏢 Áreas creadas:          {len(areas_creadas)}")
        print(f"   💳 Métodos de pago:        {len(metodos_creados)} ({sum(1 for m in metodos_creados if m.activo)} activos)")
        print(f"   🍧 Sabores disponibles:     {len(sabores_creados)}")
        print(f"   📝 Pedidos generados:       {pedidos_creados}")
        print(f"   🧾 Facturas generadas:      {len(facturas_generadas)}")
        
        # Mostrar rango de facturas
        if facturas_generadas:
            print(f"\n🧾 Rango de facturas generadas:")
            print(f"   Primera factura: {facturas_generadas[0]}")
            print(f"   Última factura:  {facturas_generadas[-1]}")
            print(f"   Total facturas:  {len(set(facturas_generadas))} (sin duplicados)")
        
        # Estadísticas de stock
        print(f"\n📈 Stock final de sabores:")
        for sabor in sabores_creados:
            vendido = sabor.stock_inicial - sabor.stock_disponible
            porcentaje = (vendido / sabor.stock_inicial) * 100 if sabor.stock_inicial > 0 else 0
            barra = "█" * int(porcentaje / 5) + "░" * (20 - int(porcentaje / 5))
            print(f"   {sabor.nombre.capitalize():15} | {barra} | {sabor.stock_disponible:3d}/{sabor.stock_inicial:3d} | {porcentaje:.0f}% vendido")
        
        # Estadísticas de pedidos
        total_usd = sum(p.cantidad * p.sabor.precio_usd for p in Pedido.query.all() if p.sabor)
        total_pagados = Pedido.query.filter_by(pago='si').count()
        total_pendientes = Pedido.query.filter_by(pago='no').count()
        
        print(f"\n💰 Resumen financiero:")
        print(f"   Total facturado (USD):    ${total_usd:.2f}")
        print(f"   Total facturado (Bs):     Bs. {total_usd * TASA_BCV_SIMULADA:.2f}")
        print(f"   Pedidos pagados:          {total_pagados} ({total_pagados/pedidos_creados*100:.0f}%)")
        print(f"   Pedidos pendientes:       {total_pendientes} ({total_pendientes/pedidos_creados*100:.0f}%)")
        
        # Pedidos por método de pago
        print(f"\n💳 Pedidos por método de pago:")
        for metodo in metodos_creados:
            count = Pedido.query.filter_by(metodo_pago_codigo=metodo.codigo, pago='si').count()
            if count > 0:
                total_metodo = sum(
                    p.cantidad * p.sabor.precio_usd 
                    for p in Pedido.query.filter_by(metodo_pago_codigo=metodo.codigo, pago='si').all() 
                    if p.sabor
                )
                print(f"   {metodo.nombre:25} | {count:2d} pedidos | ${total_metodo:.2f}")
        
        # Verificar secuencia de facturas
        print(f"\n🔍 Verificación de facturas:")
        facturas_ordenadas = sorted(facturas_generadas)
        secuencia_correcta = True
        for i in range(1, len(facturas_ordenadas)):
            num_anterior = int(facturas_ordenadas[i-1].split('-')[-1])
            num_actual = int(facturas_ordenadas[i].split('-')[-1])
            if num_actual != num_anterior + 1:
                print(f"   ⚠️  Salto en secuencia: {facturas_ordenadas[i-1]} → {facturas_ordenadas[i]}")
                secuencia_correcta = False
        
        if secuencia_correcta:
            print(f"   ✅ Secuencia de facturas correcta (incremento de 1 en 1)")
        
        print("\n" + "=" * 70)
        print("✅ ¡Base de datos poblada con éxito!")
        print("🚀 Ejecuta: python app.py")
        print("🌐 Abre: http://localhost:5000")
        print("💡 Prueba buscar facturas en: Historial > Buscar Factura")
        print("=" * 70 + "\n")

if __name__ == '__main__':
    llenar_base_datos()