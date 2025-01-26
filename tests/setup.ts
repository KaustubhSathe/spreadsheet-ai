import '@testing-library/jest-dom';

// Mock crypto for UUID generation
Object.defineProperty(window, 'crypto', {
  value: {
    randomUUID: () => '123e4567-e89b-12d3-a456-426614174000'
  }
});
