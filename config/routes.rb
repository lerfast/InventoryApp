Rails.application.routes.draw do
  # ... otras rutas ...

  get 'excel/read_excel', to: 'excel#read_excel'
  get 'excel/write_excel_form', to: 'excel#write_excel_form'
  post 'excel/write_excel', to: 'excel#write_excel'
end
