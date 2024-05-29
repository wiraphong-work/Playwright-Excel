import { Page } from '@playwright/test';

export class requestAPI {
  constructor(private page: Page) {}

  async getAllProducts() {
    const response = await this.page.request.get('/products');
    return response.json();
  }

  async getAllCategory() {
    const response = await this.page.request.get('/products/categories');
    return response.json();
  }

  async getAllCart() {
    const response = await this.page.request.get('/carts');
    return response.json();
  }
}
