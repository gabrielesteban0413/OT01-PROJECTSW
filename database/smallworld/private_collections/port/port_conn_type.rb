_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_out.txt"

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("mit_rme_port")
    mit_rme_ports << vista.collection(:mit_rme_port)

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

        id_mit_rme_port << texto[1].as_number()
        nuevo_valor << texto[2].write_string

        posibles_mit_rme_ports << mit_rme_ports.select(predicate.eq(:id, id_mit_rme_port))
        un_mit_rme_port << posibles_mit_rme_ports.an_element()

        _if un_mit_rme_port _isnt _unset
        _then
            _try
                un_mit_rme_port.type << nuevo_valor   
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

    show("--------------GS--EDIT--SPEC--------.")
    show("exitosos: ", contador_exitosos)
    show("fallidos: ", contador_fallidos)

_endblock