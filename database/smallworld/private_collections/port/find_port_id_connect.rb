_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_out.txt"
    ruta_salida << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_find.txt"


    input_file << external_text_input_stream.new(ruta_fuente)
    output_file << external_text_output_stream.new(ruta_salida)

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("Findmit_rme_port")
    mit_rme_ports << vista.collection(:mit_rme_port)
    results << rope.new()

    _for una_linea _over 1.upto(100000000)
    _loop
        linea << input_file.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        id << linea.write_string
        id_a_buscar << id.as_number()

        posibles << mit_rme_ports.select(predicate.eq(:id, id_a_buscar))
        un_objeto << posibles.an_element()

        _if un_objeto _isnt _unset
        _then
            connected_text << un_objeto.connected.write_string
            results.add( '"' + id + '"|' + connected_text )
        _else
            results.add(id + "&!")
        _endif
    _endloop

    input_file.close()

    _for r _over results.elements()
    _loop
        output_file.write(r)
        output_file.newline()
    _endloop
    output_file.close()

    show("--------GS-FIND---------")
_endblock