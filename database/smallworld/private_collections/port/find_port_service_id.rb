#DE ID A IDSERVICIO
_block
    ruta_fuente << "C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\00_out.txt"
    ruta_salida << "C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\00_find.txt"


    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("FIND_PT_CB")
    mit_rme_ports << vista.collection(:mit_rme_port)

    resultados << rope.new()

    _for una_linea _over 1.upto(100000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        # Obtener el ID como cadena
        id_cadena << linea.write_string
        # Convertir a número usando as_number() (como en el ejemplo)
        id_numero << id_cadena.as_number()

        # Buscar por el campo ID (numérico)
        posibles << mit_rme_ports.select(predicate.eq(:id, id_numero))
        un_objeto << posibles.an_element()

        _if un_objeto _isnt _unset
        _then
            # Obtener el ID del objeto (numérico) y el atributo connected
            id_texto << un_objeto.id.write_string
            # El atributo se llama "connected" (con interrogante) según errores previos
            conectado_texto << un_objeto.z.write_string
            resultados.add( id_texto + "|" + id_cadena + "|" + conectado_texto)
        _else
            resultados.add("N/A|" + id_cadena + "|?")
        _endif
    _endloop

    archivo_fuente.close()

    _for r _over resultados.elements()
    _loop
        archivo_salida.write(r)
        archivo_salida.newline()
    _endloop
    archivo_salida.close()

    show("--------GS-FIND---------")
_endblock