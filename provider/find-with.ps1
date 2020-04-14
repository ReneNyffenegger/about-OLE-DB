foreach ($provider in [System.Data.OleDb.OleDbEnumerator]::GetRootEnumerator()) {

   "{0,-30} {1,-55} {2} {3} {4,-30}" -f
      $provider['sources_name'       ] ,
      $provider['sources_description'] ,
      $provider['sources_clsid'      ] ,
    # $provider['sources_parsename'  ] ,
      $provider['sources_type'       ] ,
      $provider['sources_isparent'   ]
}
