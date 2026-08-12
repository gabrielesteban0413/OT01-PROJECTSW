_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_out.txt"
    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("aerial_route")
    aerial_routes << vista.collection(:aerial_route)

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    lista_modificada << rope.new()

    contador_exitosos << 0
    contador_fallidos << 0

    _for una_linea _over 1.upto(100000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        texto << linea.write_string.split_by("|", _true)
        _if texto.size <> 2
        _then
            lista_modificada.add(linea + "|FALLIDO")
            contador_fallidos << contador_fallidos + 1
            _continue
        _endif

        id_aerial_route << texto[1].as_number()
        owner_value << texto[2].write_string

        posibles_aerial_routes << aerial_routes.select(predicate.eq(:id, id_aerial_route))
        un_aerial_route << posibles_aerial_routes.an_element()

        _if un_aerial_route _isnt _unset
        _then
            _try
                un_aerial_route.owner << owner_value
                lista_modificada.add(linea + "|EXITOSO")
                contador_exitosos << contador_exitosos + 1
            _when error
                lista_modificada.add(linea + "|FALLIDO")
                contador_fallidos << contador_fallidos + 1
            _endtry
        _else
            lista_modificada.add(linea + "|FALLIDO")
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

    show("--------------GS--EDI--AEREAL-ROUTE--------.")
    show("exitosos: ", contador_exitosos)
    show("fallidos: ", contador_fallidos)
_endblock