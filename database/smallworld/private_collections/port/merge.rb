_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\scripts\\monitoring\\private_collections\\00_out.txt"
    #ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\src\\private_collections\\02_mit_rme_port\\edit\\P00_MERGE.txt"
    
    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("mit_rme_port_ot")
    mit_rme_ports << vista.collection(:mit_rme_port)

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    lista_modificada << rope.new()

    contador_type << 0
    contador_physical_status << 0
    contador_connector_type << 0
    contador_tipo_senial << 0
    contador_ancho_de_banda << 0
    contador_unidad_ancho_banda << 0
    contador_proyecto_red << 0
    contador_cambio_serv_man << 0
    contador_fecha_aprovisionamiento << 0
    contador_direccion_cliente_id << 0
    contador_contratista << 0
    contador_id_servicio << 0
    contador_equipo_central << 0
    contador_puerto_central << 0
    contador_tipo_servicio << 0
    contador_nombre_cliente << 0
    contador_comentarios << 0
    
    contador_fallidos << 0
    contador_total_registros << 0

    _for una_linea _over 1.upto(100000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        contador_total_registros << contador_total_registros + 1
        
        campos << linea.write_string.split_by("|", _true)
        
        _if campos.size <> 18
        _then
            lista_modificada.add(linea + "|FALLIDO_CAMPOS_INSUFICIENTES (se esperaban 17, se recibieron " + campos.size.write_string + ")")
            contador_fallidos << contador_fallidos + 1
            _continue
        _endif

        id_mit_rme_port << campos[1].as_number()
        valor_physical_status << campos[2].write_string   
        valor_type << campos[3].write_string              
        valor_connector_type << campos[4].write_string
        valor_tipo_senial << campos[5].write_string
        valor_ancho_de_banda << campos[6].as_number()
        valor_unidad_ancho_banda << campos[7].write_string
        valor_proyecto_red << campos[8].write_string
        valor_cambio_serv_man << campos[9].write_string
        valor_fecha_aprovisionamiento_str << campos[10].write_string
        valor_contratista << campos[11].write_string        
        valor_id_servicio << campos[12].write_string        
        valor_equipo_central << campos[13].write_string     
        valor_puerto_central << campos[14].write_string   
        valor_tipo_servicio << campos[15].write_string     
        valor_nombre_cliente << campos[16].write_string     
        valor_direccion_cliente_id << campos[17].as_number()
        valor_comentarios << campos[18].write_string       

        posibles_mit_rme_ports << mit_rme_ports.select(predicate.eq(:id, id_mit_rme_port))
        un_mit_rme_port << posibles_mit_rme_ports.an_element()

        _if un_mit_rme_port _isnt _unset
        _then
            exito_total << _true
            resultados_campos << rope.new()
            
            _try
                un_mit_rme_port.type << valor_type
                resultados_campos.add("type:OK")
                contador_type << contador_type + 1
            _when error
                resultados_campos.add("type:ERROR")
                exito_total << _false
            _endtry
            
            _try
                un_mit_rme_port.physical_status << valor_physical_status
                resultados_campos.add("physical_status:OK")
                contador_physical_status << contador_physical_status + 1
            _when error
                resultados_campos.add("physical_status:ERROR")
                exito_total << _false
            _endtry
            
            _try
                un_mit_rme_port.connector_type << valor_connector_type
                resultados_campos.add("connector_type:OK")
                contador_connector_type << contador_connector_type + 1
            _when error
                resultados_campos.add("connector_type:ERROR")
                exito_total << _false
            _endtry
            
            _try
                un_mit_rme_port.tipo_senial << valor_tipo_senial
                resultados_campos.add("tipo_senial:OK")
                contador_tipo_senial << contador_tipo_senial + 1
            _when error
                resultados_campos.add("tipo_senial:ERROR")
                exito_total << _false
            _endtry
            
            _try
                un_mit_rme_port.ancho_de_banda << valor_ancho_de_banda
                resultados_campos.add("ancho_de_banda:OK")
                contador_ancho_de_banda << contador_ancho_de_banda + 1
            _when error
                resultados_campos.add("ancho_de_banda:ERROR")
                exito_total << _false
            _endtry
            
            _try
                un_mit_rme_port.unidad_ancho_banda << valor_unidad_ancho_banda
                resultados_campos.add("unidad_ancho_banda:OK")
                contador_unidad_ancho_banda << contador_unidad_ancho_banda + 1
            _when error
                resultados_campos.add("unidad_ancho_banda:ERROR")
                exito_total << _false
            _endtry
            
            _try
                un_mit_rme_port.proyecto_red << valor_proyecto_red
                resultados_campos.add("proyecto_red:OK")
                contador_proyecto_red << contador_proyecto_red + 1
            _when error
                resultados_campos.add("proyecto_red:ERROR")
                exito_total << _false
            _endtry
            
            _try
                un_mit_rme_port.cambio_serv_man << valor_cambio_serv_man
                resultados_campos.add("cambio_serv_man:OK")
                contador_cambio_serv_man << contador_cambio_serv_man + 1
            _when error
                resultados_campos.add("cambio_serv_man:ERROR")
                exito_total << _false
            _endtry

            _try
                _if valor_fecha_aprovisionamiento_str.default("") _isnt ""
                _then
                    fecha_aprovisionamiento << ds_date.new_from_string(valor_fecha_aprovisionamiento_str)
                    un_mit_rme_port.fecha_aprovisionamiento << fecha_aprovisionamiento
                _else
                    un_mit_rme_port.fecha_aprovisionamiento << _unset
                _endif
                resultados_campos.add("fecha_aprovisionamiento:OK")
                contador_fecha_aprovisionamiento << contador_fecha_aprovisionamiento + 1
            _when error
                resultados_campos.add("fecha_aprovisionamiento:ERROR")
                exito_total << _false
            _endtry
            
            _try
                un_mit_rme_port.direccion_cliente_id << valor_direccion_cliente_id
                resultados_campos.add("direccion_cliente_id:OK")
                contador_direccion_cliente_id << contador_direccion_cliente_id + 1
            _when error
                resultados_campos.add("direccion_cliente_id:ERROR")
                exito_total << _false
            _endtry

            _try
                un_mit_rme_port.contratista << valor_contratista
                resultados_campos.add("contratista:OK")
                contador_contratista << contador_contratista + 1
            _when error
                resultados_campos.add("contratista:ERROR")
                exito_total << _false
            _endtry

            _try
                un_mit_rme_port.id_servicio << valor_id_servicio
                resultados_campos.add("id_servicio:OK")
                contador_id_servicio << contador_id_servicio + 1
            _when error
                resultados_campos.add("id_servicio:ERROR")
                exito_total << _false
            _endtry

            _try
                un_mit_rme_port.equipo_central << valor_equipo_central
                resultados_campos.add("equipo_central:OK")
                contador_equipo_central << contador_equipo_central + 1
            _when error
                resultados_campos.add("equipo_central:ERROR")
                exito_total << _false
            _endtry

            _try
                un_mit_rme_port.puerto_central << valor_puerto_central
                resultados_campos.add("puerto_central:OK")
                contador_puerto_central << contador_puerto_central + 1
            _when error
                resultados_campos.add("puerto_central:ERROR")
                exito_total << _false
            _endtry

            _try
                un_mit_rme_port.tipo_servicio << valor_tipo_servicio
                resultados_campos.add("tipo_servicio:OK")
                contador_tipo_servicio << contador_tipo_servicio + 1
            _when error
                resultados_campos.add("tipo_servicio:ERROR")
                exito_total << _false
            _endtry

            _try
                un_mit_rme_port.nombre_cliente << valor_nombre_cliente
                resultados_campos.add("nombre_cliente:OK")
                contador_nombre_cliente << contador_nombre_cliente + 1
            _when error
                resultados_campos.add("nombre_cliente:ERROR")
                exito_total << _false
            _endtry

            _try
                un_mit_rme_port.comentarios << valor_comentarios
                resultados_campos.add("comentarios:OK")
                contador_comentarios << contador_comentarios + 1
            _when error
                resultados_campos.add("comentarios:ERROR")
                exito_total << _false
            _endtry

            _if exito_total
            _then
                lista_modificada.add(linea + "|GS_EXITOSOS")
            _else
                resultados_str << ""
                primero << _true
                _for elem _over resultados_campos.elements()
                _loop
                    _if primero
                    _then
                        primero << _false
                    _else
                        resultados_str << resultados_str + ", "
                    _endif
                    resultados_str << resultados_str + elem
                _endloop
                lista_modificada.add(linea + "|CAMPOS_PARCIALES|" + resultados_str)
                contador_fallidos << contador_fallidos + 1
            _endif
        _else
            lista_modificada.add(linea + "|ID_NO_ENCONTRADO")
            contador_fallidos << contador_fallidos + 1
        _endif
    _endloop

    archivo_fuente.close()
    vista.commit()

    archivo_fuente << external_text_output_stream.new(ruta_fuente)
    _for nueva_linea _over lista_modificada.elements()
    _loop
        archivo_fuente.write(nueva_linea)
        archivo_fuente.newline()
    _endloop
    archivo_fuente.close()

    show("TOTAL =", contador_total_registros)
    show("Registros con errores: ", contador_fallidos)
    show("------------GS-MERGE-PORT--------------------")
_endblock