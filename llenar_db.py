import random
from datetime import datetime, timedelta
from app import app, db, Sabor, Area, Pedido, MetodoPago

# Datos simulados para llenar la base de datos
NOMBRES = [
    "Carlos Rodríguez", "Lesgleiny Martínez", "Mayerlin González", 
    "Emiro Pérez", "Cordero José", "Zoraida Hernández", 
    "Raquel Díaz", "Abigail Torres", "Juan Morales", 
    "María Fernández", "Pedro Ramírez", "Luis Castillo",
    "Ana Suárez", "Jorge Mendoza", "Carmen Ortega", "Rafael Vargas"
]

AREAS_DATA = [
    "ALMACEN", 
    "SISTEMAS", 
    "COSMETICOS", 
    "FARMACIA", 
    "RECURSOS HUMANOS", 
    "CAJA",
    "ADMINISTRACIÓN",
    "VENTAS"
]

# Definimos los sabores con su stock inicial y un precio asignado en USD
SABORES_DATA = [
    ("fresa", 50, 1.0), 
    ("colita", 60, 1.0), 
    ("crema reina", 40, 1.5), 
    ("limon", 30, 1.0), 
    ("mantecado", 45, 1.5),
    ("uva", 35, 2.0),
    ("chocolate", 25, 2.0),
    ("coco", 20, 1.5)
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
        'activo': False  # Lo creamos inactivo para probar la funcionalidad
    }
]

# Tasa BCV simulada para los cálculos
TASA_BCV_SIMULADA = 530.50

def generar_monto_bs(monto_usd, metodo_codigo):
    """Genera un monto en Bs simulado según el método de pago"""
    if metodo_codigo in ['efectivo_bs', 'pago_movil', 'transferencia', 'punto_venta']:
        # A veces se paga exacto, a veces con propina
        if random.random() < 0.3:  # 30% de probabilidad de incluir propina
            propina = random.uniform(0.5, 2.0)
            return round((monto_usd * TASA_BCV_SIMULADA) + propina, 2)
        else:
            return round(monto_usd * TASA_BCV_SIMULADA, 2)
    return None

def llenar_base_datos():
    with app.app_context():
        print("🗑️  Limpiando base de datos anterior...")
        db.drop_all()  # Borra todas las tablas existentes
        db.create_all() # Las vuelve a crear limpias con la nueva estructura

        print("\n🏢 Creando Áreas Organizacionales...")
        areas_creadas = []
        for nombre_area in AREAS_DATA:
            nueva_area = Area(nombre=nombre_area)
            db.session.add(nueva_area)
            areas_creadas.append(nueva_area)
            print(f"   ✓ Área: {nombre_area}")
        
        db.session.commit()

        print("\n💳 Creando Métodos de Pago...")
        metodos_creados = []
        for metodo_data in METODOS_PAGO_DATA:
            nuevo_metodo = MetodoPago(**metodo_data)
            db.session.add(nuevo_metodo)
            metodos_creados.append(nuevo_metodo)
            estado = "ACTIVO" if metodo_data['activo'] else "INACTIVO"
            print(f"   ✓ Método: {metodo_data['nombre']} ({metodo_data['codigo']}) - {estado}")
            print(f"     - Requiere Capture: {'Sí' if metodo_data['requiere_capture'] else 'No'}")
            print(f"     - Requiere Monto Bs: {'Sí' if metodo_data['requiere_monto_bs'] else 'No'}")
        
        db.session.commit()

        print("\n🍧 Creando Sabores, Precios y Stock inicial...")
        sabores_creados = []
        for nombre_sabor, stock, precio in SABORES_DATA:
            nuevo_sabor = Sabor(
                nombre=nombre_sabor, 
                stock_inicial=stock, 
                stock_disponible=stock,
                precio_usd=precio
            )
            db.session.add(nuevo_sabor)
            sabores_creados.append(nuevo_sabor)
            print(f"   ✓ Sabor: {nombre_sabor.capitalize()} - Stock: {stock} - ${precio:.2f}")
            
        db.session.commit()

        print("\n📝 Generando 50 pedidos aleatorios con datos completos...")
        pedidos_creados = 0
        fecha_base = datetime.now() - timedelta(days=7)  # Pedidos de la última semana
        
        for i in range(50):
            nombre = random.choice(NOMBRES)
            area = random.choice(areas_creadas)
            sabor = random.choice(sabores_creados)
            cantidad = random.randint(1, 4)  # Pedidos de 1 a 4 cepillados
            pago = random.choice(["si", "si", "si", "no"])  # 75% probabilidad de pagado
            entrega = random.choice(["si", "si", "no"])  # 66% probabilidad de entregado
            
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
                    
                    # Simular capture si el método lo requiere
                    if metodo_elegido.requiere_capture:
                        # 80% de probabilidad de tener capture
                        if random.random() < 0.8:
                            capture_img = f"capture_simulado_{i+1}.jpg"

            # Generar fecha aleatoria en los últimos 7 días
            dias_atras = random.randint(0, 7)
            horas_atras = random.randint(0, 23)
            minutos_atras = random.randint(0, 59)
            fecha_pedido = fecha_base + timedelta(days=dias_atras, hours=horas_atras, minutes=minutos_atras)

            # Descontar del stock si hay disponibilidad
            if sabor.stock_disponible >= cantidad:
                sabor.stock_disponible -= cantidad
                
                nuevo_pedido = Pedido(
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
                
                # Mostrar resumen del pedido creado
                estado_pago = "PAGADO" if pago == "si" else "PENDIENTE"
                metodo_str = metodo_pago_codigo.replace('_', ' ').title() if pago == "si" else "N/A"
                print(f"   #{pedidos_creados} {nombre} - {sabor.nombre.capitalize()} x{cantidad} - ${cantidad * sabor.precio_usd:.2f} - {estado_pago} ({metodo_str}) - {fecha_pedido.strftime('%d/%m/%Y %H:%M')}")

        db.session.commit()
        
        # Mostrar resumen final
        print("\n" + "="*60)
        print("📊 RESUMEN DE LA BASE DE DATOS GENERADA")
        print("="*60)
        print(f"🏢 Áreas creadas: {len(areas_creadas)}")
        print(f"💳 Métodos de pago: {len(metodos_creados)}")
        print(f"🍧 Sabores disponibles: {len(sabores_creados)}")
        print(f"📝 Pedidos generados: {pedidos_creados}")
        
        # Calcular estadísticas
        pedidos_pagados = sum(1 for _ in range(pedidos_creados) if True)  # Placeholder
        total_usd = 0
        total_bs = 0
        
        print("\n📈 Stock final de sabores:")
        for sabor in sabores_creados:
            vendido = sabor.stock_inicial - sabor.stock_disponible
            print(f"   • {sabor.nombre.capitalize()}: {sabor.stock_disponible}/{sabor.stock_inicial} disponibles (Vendidos: {vendido})")
        
        print("\n✅ ¡Base de datos poblada con éxito!")
        print("🚀 Ya puedes iniciar tu app.py para probar el sistema")
        print("="*60 + "\n")

if __name__ == '__main__':
    llenar_base_datos()