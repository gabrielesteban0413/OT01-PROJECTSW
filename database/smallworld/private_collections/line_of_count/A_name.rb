_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_out.txt"

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("FI_OLDNAME_HILO")

    copper_line_of_counts << vista.collection(:copper_line_of_count)

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    lista_modificada << rope.new()
    contador_exitosos << 0
    contador_fallidos << 0

    _for una_linea _over 1.upto(10000)
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

        id_hilo << texto[1].as_number()
        nombre_nuevo_value << texto[2].write_string

        posibles_hilos << copper_line_of_counts.select(predicate.eq(:id, id_hilo))
        un_hilo << posibles_hilos.an_element()

        _if un_hilo _isnt _unset
        _then
            _try
                un_hilo.designation << nombre_nuevo_value
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

    show("--------------GS--EDIT--NOMBRE--HILO--------.")
    show("exitosos: ", contador_exitosos)
    show("fallidos: ", contador_fallidos)
_endblock