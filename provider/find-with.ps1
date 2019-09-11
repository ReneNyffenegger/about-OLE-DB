foreach ($provider in [System.Data.OleDb.OleDbEnumerator]::GetRootEnumerator()) {

   write-output($provider['sources_name'       ]  + "`t" + 
                $provider['sources_description']  + "`t" +
                $provider['sources_clsid'      ]  + "`t" +
   #            $provider['sources_parsename'  ]  + "`t" +
                $provider['sources_type'       ]  + "`t" +
                $provider['sources_isparent'   ] 
   )

}
