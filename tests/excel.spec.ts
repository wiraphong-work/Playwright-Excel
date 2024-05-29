import { test } from '@playwright/test';
import { requestAPI } from './request-api';
import ExcelJS from 'exceljs';
import styleExcel from './style.json';

test('table excel', async ({ page }) => {
  const workbook = new ExcelJS.Workbook();
  const request = new requestAPI(page);

  const worksheetProdutcs = await workbook.addWorksheet('Get All Products');
  const worksheetCategorys = await workbook.addWorksheet('Get All Categorys');
  const worksheetCarts = await workbook.addWorksheet('Get All Carts');

  //Export Excel Sheet Product
  const responseProduct = await request.getAllProducts();
  worksheetProdutcs.addTable({
    name: 'responseProduct',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      showRowStripes: true,
    },
    columns: styleExcel.col_header_product,
    rows: responseProduct.map((products, index) => {
      return [
        index + 1,
        products.id,
        products.title,
        products.price,
        products.description,
        products.category,
        products.image,
        products.rating.rate,
        products.rating.count,
      ];
    }),
  });
  worksheetProdutcs.columns = styleExcel.col_style_product;

//Export Excel Sheet Category
  const responseCategory = await request.getAllCategory();
  worksheetCategorys.addTable({
    name: 'responseCategory',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      showRowStripes: true,
    },
    columns: styleExcel.col_header_category,
    rows: responseCategory.map((category, index) => {
      return [index + 1, category];
    }),
  });
  worksheetCategorys.columns = styleExcel.col_style_category;

  //Export Excel Sheet Carts
  const responseCarts = await request.getAllCart();
  const aphabet = 'ABCD'.split('');

  aphabet.map((val) => {
    worksheetCarts.getCell(`${val}1`).alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };
  });

  worksheetCarts.addTable({
    name: 'responseCarts',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      showRowStripes: true,
    },
    columns: styleExcel.col_header_cart,
    rows: responseCarts.map((cart, index) => {
      const value = cart.products.map((val) => {
        return `ProductId: '${val.productId}' , Quantity: '${val.quantity}'`;
      });
      worksheetCarts.getCell(`A${index + 2}`).alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
      worksheetCarts.getCell(`B${index + 2}`).alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
      worksheetCarts.getCell(`C${index + 2}`).alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
      worksheetCarts.getCell(`D${index + 2}`).alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
      worksheetCarts.getCell(`E${index + 2}`).alignment = { wrapText: true };
      return [index + 1, cart.id, cart.userId, cart.date, value.join('\n')];
    }),
  });
  worksheetCarts.columns = styleExcel.col_style_cart;

  await workbook.xlsx.writeFile('./excel/output.xlsx');
});
